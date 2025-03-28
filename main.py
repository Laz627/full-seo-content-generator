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
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from io import BytesIO
import base64
import random
from typing import List, Dict, Any, Tuple, Optional
import logging
import traceback
import openpyxl
import matplotlib.pyplot as plt
import altair as alt

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
                                for expanded in paa_item.get('expanded_element', []):
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
                            for expanded in item.get('expanded_element', []):
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
            "include_serp_info": True,
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
                logger.warning(f"Invalid API response: {data.get('status_message')}")
                return [], False
            
            keyword_suggestions = []
            
            # Process results based on the specific JSON structure from sample
            for task in data['tasks']:
                if not task.get('result'):
                    continue
                
                for result in task['result']:
                    # First, check for seed keyword data
                    if 'seed_keyword_data' in result and result['seed_keyword_data']:
                        seed_data = result['seed_keyword_data']
                        if 'keyword_info' in seed_data:
                            keyword_info = seed_data['keyword_info']
                            keyword_suggestions.append({
                                'keyword': result.get('seed_keyword', ''),
                                'search_volume': keyword_info.get('search_volume', 0),
                                'cpc': keyword_info.get('cpc', 0.0),
                                'competition': keyword_info.get('competition', 0.0)
                            })
                    
                    # Then look for items array which contains related keywords
                    if 'items' in result and isinstance(result['items'], list):
                        for item in result['items']:
                            if 'keyword_info' in item:
                                keyword_info = item['keyword_info']
                                keyword_suggestions.append({
                                    'keyword': item.get('keyword', ''),
                                    'search_volume': keyword_info.get('search_volume', 0),
                                    'cpc': keyword_info.get('cpc', 0.0),
                                    'competition': keyword_info.get('competition', 0.0)
                                })
            
            # Check if we successfully found keywords
            if keyword_suggestions:
                # Sort by search volume (descending)
                keyword_suggestions.sort(key=lambda x: x.get('search_volume', 0), reverse=True)
                logger.info(f"Successfully extracted {len(keyword_suggestions)} keyword suggestions")
                return keyword_suggestions, True
            else:
                logger.warning(f"No keyword suggestions found in the response")
                return [], False
        else:
            error_msg = f"HTTP Error: {response.status_code} - {response.text}"
            logger.error(error_msg)
            return [], False
    
    except Exception as e:
        error_msg = f"Exception in fetch_keyword_suggestions: {str(e)}"
        logger.error(error_msg)
        logger.error(traceback.format_exc())
        return [], False

def fetch_related_keywords_dataforseo(keyword: str, api_login: str, api_password: str) -> Tuple[List[Dict], bool]:
    """
    Fetch related keywords from DataForSEO Related Keywords API
    Uses proper API structure parsing based on the provided sample
    Returns: related_keywords, success_status
    """
    try:
        url = "https://api.dataforseo.com/v3/dataforseo_labs/google/related_keywords/live"
        headers = {
            'Content-Type': 'application/json',
        }
        
        # Prepare request data based on the provided sample format
        post_data = [{
            "keyword": keyword,
            "language_name": "English", 
            "location_code": 2840,  # USA
            "limit": 20  # Fetch top 20 related keywords
        }]
        
        # Log the request for debugging
        logger.info(f"Fetching related keywords for: {keyword}")
        
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
            logger.info(f"API Response status: {data.get('status_code')}")
            
            # Validate response
            if data.get('status_code') != 20000 or not data.get('tasks'):
                logger.warning(f"Invalid API response: {data.get('status_message')}")
                return fetch_keyword_suggestions(keyword, api_login, api_password)
            
            related_keywords = []
            
            # Process the results following the structure of the provided JSON example
            for task in data['tasks']:
                if not task.get('result'):
                    continue
                
                for result in task['result']:
                    # Process items array which contains the keyword data
                    for item in result.get('items', []):
                        if 'keyword_data' in item:
                            kw_data = item['keyword_data']
                            if 'keyword_info' in kw_data:
                                keyword_info = kw_data['keyword_info']
                                
                                related_keywords.append({
                                    'keyword': kw_data.get('keyword', ''),
                                    'search_volume': keyword_info.get('search_volume', 0),
                                    'cpc': keyword_info.get('cpc', 0.0),
                                    'competition': keyword_info.get('competition', 0.0)
                                })
            
            # Check if we found any keywords with this approach
            if related_keywords:
                logger.info(f"Successfully extracted {len(related_keywords)} related keywords")
                return related_keywords, True
            else:
                # If no keywords found, try the keyword suggestions endpoint
                logger.warning(f"No related keywords found, trying keyword suggestions endpoint")
                return fetch_keyword_suggestions(keyword, api_login, api_password)
        else:
            error_msg = f"HTTP Error: {response.status_code} - {response.text}"
            logger.error(error_msg)
            return fetch_keyword_suggestions(keyword, api_login, api_password)
    
    except Exception as e:
        error_msg = f"Exception in fetch_related_keywords_dataforseo: {str(e)}"
        logger.error(error_msg)
        logger.error(traceback.format_exc())
        return fetch_keyword_suggestions(keyword, api_login, api_password)

###############################################################################
# 4. Web Scraping and Content Analysis
###############################################################################

def scrape_webpage(url: str) -> Tuple[str, bool]:
    """
    Enhanced webpage scraping with better error handling
    Returns: content, success_status
    """
    try:
        # Use trafilatura without headers parameter
        downloaded = trafilatura.fetch_url(url)
        if downloaded:
            content = trafilatura.extract(downloaded, include_comments=False, include_tables=True)
            if content:
                return content, True
        
        # Fallback to requests + BeautifulSoup if trafilatura fails
        user_agents = [
            'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
            'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/14.1.1 Safari/605.1.15',
            'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:89.0) Gecko/20100101 Firefox/89.0',
            'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/92.0.4515.107 Safari/537.36'
        ]
        
        headers = {
            'User-Agent': random.choice(user_agents),
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
            'Accept-Language': 'en-US,en;q=0.5',
            'Referer': 'https://www.google.com/'
        }
        
        response = requests.get(url, headers=headers, timeout=15)
        
        if response.status_code == 200:
            soup = BeautifulSoup(response.text, 'html.parser')
            
            # Remove script, style, and other non-content elements
            for element in soup(["script", "style", "header", "footer", "nav", "aside", "form"]):
                element.extract()
            
            # Try to find main content
            main_content = soup.find('main') or soup.find('article') or soup.find('div', class_='content')
            
            if main_content:
                text = main_content.get_text(separator='\n')
            else:
                text = soup.get_text(separator='\n')
            
            # Clean up text
            lines = (line.strip() for line in text.splitlines())
            chunks = (phrase.strip() for line in lines for phrase in line.split("  "))
            text = '\n'.join(chunk for chunk in chunks if chunk)
            
            return text, True
        
        elif response.status_code == 403:
            logger.warning(f"Access forbidden (403) for URL: {url}")
            return f"[Content not accessible due to site restrictions]", False
        
        else:
            logger.warning(f"HTTP error {response.status_code} for URL: {url}")
            return f"[Error retrieving content: HTTP {response.status_code}]", False
    
    except Exception as e:
        error_msg = f"Exception in scrape_webpage for {url}: {str(e)}"
        logger.error(error_msg)
        return f"[Error: {str(e)}]", False

def extract_headings(url: str) -> Dict[str, List[str]]:
    """
    Extract headings (H1, H2, H3) from a webpage
    """
    try:
        # List of common user agents to rotate
        user_agents = [
            'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
            'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/14.1.1 Safari/605.1.15',
            'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:89.0) Gecko/20100101 Firefox/89.0',
            'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/92.0.4515.107 Safari/537.36'
        ]
        
        # Select a random user agent
        user_agent = random.choice(user_agents)
        
        headers = {
            'User-Agent': user_agent,
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
            'Accept-Language': 'en-US,en;q=0.5',
            'Referer': 'https://www.google.com/'
        }
        
        response = requests.get(url, headers=headers, timeout=15)
        
        if response.status_code == 200:
            soup = BeautifulSoup(response.text, 'html.parser')
            
            headings = {
                'h1': [h.get_text().strip() for h in soup.find_all('h1')],
                'h2': [h.get_text().strip() for h in soup.find_all('h2')],
                'h3': [h.get_text().strip() for h in soup.find_all('h3')]
            }
            
            return headings
        else:
            return {'h1': [], 'h2': [], 'h3': []}
    
    except Exception as e:
        error_msg = f"Exception in extract_headings: {str(e)}"
        logger.error(error_msg)
        return {'h1': [], 'h2': [], 'h3': []}

###############################################################################
# 5. Content Scoring Functions
###############################################################################

def extract_important_terms(competitor_contents: List[Dict], anthropic_api_key: str) -> Tuple[Dict, bool]:
    """
    Extract important terms and topics from competitor content using Claude 3.7 Sonnet
    Returns: term_data, success_status
    """
    try:
        client = anthropic.Anthropic(api_key=anthropic_api_key)
        
        # Combine all content for analysis
        combined_content = "\n\n".join([c.get('content', '') for c in competitor_contents if c.get('content')])
        
        # Prepare summarized content if it's too long
        if len(combined_content) > 10000:
            combined_content = combined_content[:10000]
        
        # Use Claude to analyze content and extract important terms
        response = client.messages.create(
            model="claude-3-7-sonnet-20250219",
            max_tokens=1500,
            system="You are an SEO expert specializing in content analysis.",
            messages=[
                {"role": "user", "content": f"""
                Analyze the following content from top-ranking pages and extract:
                
                1. Primary terms (most important for this topic, maximum 15 terms)
                2. Secondary terms (supporting terms for this topic, maximum 20 terms)
                3. Questions being answered (maximum 10 questions)
                4. Topics that need to be covered (maximum 10 topics)
                
                Format your response as JSON:
                {{
                    "primary_terms": [
                        {{"term": "term1", "importance": 0.95, "recommended_usage": 5}},
                        {{"term": "term2", "importance": 0.85, "recommended_usage": 3}},
                        ...
                    ],
                    "secondary_terms": [
                        {{"term": "term1", "importance": 0.75, "recommended_usage": 2}},
                        {{"term": "term2", "importance": 0.65, "recommended_usage": 1}},
                        ...
                    ],
                    "questions": [
                        "Question 1?",
                        "Question 2?",
                        ...
                    ],
                    "topics": [
                        {{"topic": "Topic 1", "description": "This topic covers..."}},
                        {{"topic": "Topic 2", "description": "This topic covers..."}},
                        ...
                    ]
                }}
                
                Content to analyze:
                {combined_content}
                """}
            ],
            temperature=0.3
        )
        
        # Extract and parse JSON response
        content = response.content[0].text
        json_match = re.search(r'({.*})', content, re.DOTALL)
        if json_match:
            content = json_match.group(1)
        
        term_data = json.loads(content)
        return term_data, True
    
    except Exception as e:
        error_msg = f"Exception in extract_important_terms: {str(e)}"
        logger.error(error_msg)
        return {}, False

def score_content(content: str, term_data: Dict, keyword: str) -> Tuple[Dict, bool]:
    """
    Score content based on keyword usage, semantic relevance, and comprehensiveness
    Returns: score_data, success_status
    """
    try:
        # Initialize scores
        keyword_score = 0
        primary_terms_score = 0
        secondary_terms_score = 0
        topic_coverage_score = 0
        question_coverage_score = 0
        
        # Score primary keyword usage - IMPROVED SCORING
        keyword_count = len(re.findall(r'\b' + re.escape(keyword.lower()) + r'\b', content.lower()))
        word_count = len(re.findall(r'\b\w+\b', content))
        optimal_keyword_count = max(2, min(10, int(word_count * 0.01)))  # Between 0.5% and 2%

        # Calculate keyword score (max 100)
        if keyword_count > 0:
            if keyword_count <= optimal_keyword_count:
                # Perfect score when usage is at or below optimal
                keyword_score = (keyword_count / optimal_keyword_count) * 100
            else:
                # More generous penalty for overuse - starts penalizing after 1.5x optimal
                excessive_ratio = keyword_count / optimal_keyword_count
                if excessive_ratio <= 1.5:  # Allow up to 150% of optimal without penalty
                    keyword_score = 100
                else:
                    # Scale penalty more gradually
                    penalty = (excessive_ratio - 1.5) * 25  # Only 25% reduction per 100% over the 1.5x threshold
                    keyword_score = max(60, 100 - penalty)  # Floor of 60%
        
        # Score primary terms
        primary_term_found = 0
        primary_terms_total = len(term_data.get('primary_terms', []))
        primary_term_counts = {}
        
        if primary_terms_total > 0:
            for term_data_item in term_data.get('primary_terms', []):
                term = term_data_item.get('term', '')
                if term:
                    term_count = len(re.findall(r'\b' + re.escape(term.lower()) + r'\b', content.lower()))
                    primary_term_counts[term] = {
                        'count': term_count,
                        'importance': term_data_item.get('importance', 0),
                        'recommended': term_data_item.get('recommended_usage', 1)
                    }
                    
                    if term_count > 0:
                        primary_term_found += 1
                        
                        # Bonus for optimal usage
                        if term_count >= term_data_item.get('recommended_usage', 1):
                            if term_count <= term_data_item.get('recommended_usage', 1) * 2:
                                primary_term_found += 0.2
                    
            primary_terms_score = (primary_term_found / primary_terms_total) * 100
            
        # Score secondary terms
        secondary_term_found = 0
        secondary_terms_total = len(term_data.get('secondary_terms', []))
        secondary_term_counts = {}
        
        if secondary_terms_total > 0:
            for term_data_item in term_data.get('secondary_terms', []):
                term = term_data_item.get('term', '')
                if term:
                    term_count = len(re.findall(r'\b' + re.escape(term.lower()) + r'\b', content.lower()))
                    secondary_term_counts[term] = {
                        'count': term_count,
                        'importance': term_data_item.get('importance', 0),
                        'recommended': term_data_item.get('recommended_usage', 1)
                    }
                    
                    if term_count > 0:
                        secondary_term_found += 1
                        
            secondary_terms_score = (secondary_term_found / secondary_terms_total) * 100
        
        # Score topic coverage - IMPROVED SEMANTIC DETECTION
        topics_covered = 0
        topics_total = len(term_data.get('topics', []))
        topic_coverage = {}
        
        if topics_total > 0:
            content_lower = content.lower()
            
            for topic_data in term_data.get('topics', []):
                topic = topic_data.get('topic', '')
                description = topic_data.get('description', '')
                
                if topic:
                    # Extract key terms from both topic and its description
                    topic_terms = set(re.findall(r'\b\w{3,}\b', topic.lower()))
                    if description:
                        desc_terms = set(re.findall(r'\b\w{3,}\b', description.lower()))
                        # Combine important terms
                        key_terms = topic_terms.union(desc_terms)
                    else:
                        key_terms = topic_terms
                    
                    # Remove common words that aren't meaningful for matching
                    common_words = {'the', 'and', 'for', 'that', 'with', 'this', 'what', 'how', 
                                   'why', 'when', 'where', 'will', 'can', 'your', 'you', 'these',
                                   'those', 'them', 'they', 'some', 'have', 'has', 'had', 'are',
                                   'our', 'their', 'were', 'was', 'not', 'from', 'about'}
                    
                    key_terms = {term for term in key_terms if term not in common_words and len(term) > 2}
                    
                    # Calculate what percentage of key terms are found
                    if key_terms:
                        found_terms = sum(1 for term in key_terms if term in content_lower)
                        match_ratio = found_terms / len(key_terms)
                        
                        # Consider covered if enough key terms are found or exact match
                        is_covered = match_ratio >= 0.4 or topic.lower() in content_lower
                    else:
                        # Fall back to exact matching if no key terms
                        is_covered = topic.lower() in content_lower
                        match_ratio = 1.0 if is_covered else 0.0
                    
                    topic_coverage[topic] = {
                        'covered': is_covered,
                        'match_ratio': match_ratio,
                        'description': description,
                        'terms_found': match_ratio if key_terms else 0
                    }
                    
                    if is_covered:
                        topics_covered += 1
            
            topic_coverage_score = (topics_covered / topics_total) * 100
        
        # Score question coverage
        questions_answered = 0
        questions_total = len(term_data.get('questions', []))
        question_coverage = {}
        
        if questions_total > 0:
            for question in term_data.get('questions', []):
                core_question = question.replace('?', '').lower()
                
                # Look for core keywords from the question
                question_words = re.findall(r'\b\w+\b', core_question)
                significant_words = [w for w in question_words if len(w) > 3 and w not in ['what', 'when', 'where', 'which', 'who', 'why', 'how']]
                
                # Count how many significant words appear
                matches = sum(1 for word in significant_words if word in content.lower())
                match_ratio = 0
                if significant_words:
                    match_ratio = matches / len(significant_words)
                
                # If most significant words appear, consider the question answered
                is_answered = match_ratio >= 0.7
                question_coverage[question] = {
                    'answered': is_answered,
                    'match_ratio': match_ratio
                }
                
                if is_answered:
                    questions_answered += 1
            
            question_coverage_score = (questions_answered / questions_total) * 100
        
        # Calculate overall score (weighted)
        overall_score = (
            keyword_score * 0.2 +
            primary_terms_score * 0.35 +
            secondary_terms_score * 0.15 +
            topic_coverage_score * 0.2 +
            question_coverage_score * 0.1
        )
        
        # Compile detailed results
        score_data = {
            'overall_score': round(overall_score),
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
                'optimal_keyword_count': optimal_keyword_count,
                'primary_terms_found': primary_term_found,
                'primary_terms_total': primary_terms_total,
                'primary_term_counts': primary_term_counts,
                'secondary_terms_found': secondary_term_found,
                'secondary_terms_total': secondary_terms_total,
                'secondary_term_counts': secondary_term_counts,
                'topics_covered': topics_covered,
                'topics_total': topics_total,
                'topic_coverage': topic_coverage,
                'questions_answered': questions_answered,
                'questions_total': questions_total,
                'question_coverage': question_coverage
            },
            'grade': get_score_grade(overall_score)
        }
        
        return score_data, True
    
    except Exception as e:
        error_msg = f"Exception in score_content: {str(e)}"
        logger.error(error_msg)
        return {'overall_score': 0}, False

def get_score_grade(score: float) -> str:
    """Convert numeric score to letter grade"""
    if score >= 90:
        return "A+"
    elif score >= 85:
        return "A"
    elif score >= 80:
        return "A-"
    elif score >= 75:
        return "B+"
    elif score >= 70:
        return "B"
    elif score >= 65:
        return "B-"
    elif score >= 60:
        return "C+"
    elif score >= 55:
        return "C"
    elif score >= 50:
        return "C-"
    elif score >= 40:
        return "D"
    else:
        return "F"

def highlight_keywords_in_content(content: str, term_data: Dict, keyword: str) -> Tuple[str, bool]:
    """
    Highlight primary and secondary keywords in content with different colors
    Returns: highlighted_html, success_status
    """
    try:
        # Create a copy of the content for highlighting
        highlighted_content = content
        
        # Function to wrap term with color
        def wrap_with_span(match, color):
            term = match.group(0)
            return f'<span style="background-color: {color};">{term}</span>'
        
        # Highlight primary keyword
        pattern = re.compile(r'\b' + re.escape(keyword) + r'\b', re.IGNORECASE)
        highlighted_content = pattern.sub(lambda m: wrap_with_span(m, "#FFEB9C"), highlighted_content)
        
        # Highlight primary terms
        for term_info in term_data.get('primary_terms', []):
            term = term_info.get('term', '')
            if term and term.lower() != keyword.lower():
                pattern = re.compile(r'\b' + re.escape(term) + r'\b', re.IGNORECASE)
                highlighted_content = pattern.sub(lambda m: wrap_with_span(m, "#CDFFD8"), highlighted_content)
        
        # Highlight secondary terms
        for term_info in term_data.get('secondary_terms', []):
            term = term_info.get('term', '')
            if term:
                pattern = re.compile(r'\b' + re.escape(term) + r'\b', re.IGNORECASE)
                highlighted_content = pattern.sub(lambda m: wrap_with_span(m, "#E6F3FF"), highlighted_content)
        
        return highlighted_content, True
    
    except Exception as e:
        error_msg = f"Exception in highlight_keywords_in_content: {str(e)}"
        logger.error(error_msg)
        return content, False

def get_content_improvement_suggestions(content: str, term_data: Dict, score_data: Dict, keyword: str) -> Tuple[Dict, bool]:
    """
    Generate suggestions for improving content based on scoring results
    Returns: suggestions, success_status
    """
    try:
        suggestions = {
            'missing_terms': [],
            'underused_terms': [],
            'missing_topics': [],
            'partial_topics': [],  # New category for partially covered topics
            'unanswered_questions': [],
            'readability_suggestions': [],
            'structure_suggestions': []
        }
        
        # Check for missing primary terms
        for term_info in term_data.get('primary_terms', []):
            term = term_info.get('term', '')
            importance = term_info.get('importance', 0)
            recommended_usage = term_info.get('recommended_usage', 1)
            
            if term:
                term_count = len(re.findall(r'\b' + re.escape(term.lower()) + r'\b', content.lower()))
                
                if term_count == 0:
                    suggestions['missing_terms'].append({
                        'term': term,
                        'importance': importance,
                        'recommended_usage': recommended_usage,
                        'type': 'primary',
                        'current_usage': 0
                    })
                elif term_count < recommended_usage:
                    suggestions['underused_terms'].append({
                        'term': term,
                        'importance': importance,
                        'recommended_usage': recommended_usage,
                        'current_usage': term_count,
                        'type': 'primary'
                    })
        
        # Check for missing secondary terms (only list important ones)
        for term_info in term_data.get('secondary_terms', [])[:15]:  # Limit to top 15
            term = term_info.get('term', '')
            importance = term_info.get('importance', 0)
            recommended_usage = term_info.get('recommended_usage', 1)
            
            if term and importance > 0.5:  # Only suggest important secondary terms
                term_count = len(re.findall(r'\b' + re.escape(term.lower()) + r'\b', content.lower()))
                
                if term_count == 0:
                    suggestions['missing_terms'].append({
                        'term': term,
                        'importance': importance,
                        'recommended_usage': recommended_usage,
                        'type': 'secondary',
                        'current_usage': 0
                    })
        
        # Check for missing or partially covered topics
        topic_coverage = score_data.get('details', {}).get('topic_coverage', {})
        
        for topic_info in term_data.get('topics', []):
            topic = topic_info.get('topic', '')
            description = topic_info.get('description', '')
            
            if topic:
                if topic in topic_coverage:
                    # Get coverage data from score results
                    coverage_info = topic_coverage[topic]
                    is_covered = coverage_info.get('covered', False)
                    match_ratio = coverage_info.get('match_ratio', 0)
                    
                    # Fully covered topics (70%+ match) need no suggestions
                    if match_ratio >= 0.7:
                        continue
                    # Partially covered topics (40-69% match) need enhancement
                    elif is_covered:
                        suggestions['partial_topics'].append({
                            'topic': topic,
                            'description': description,
                            'match_ratio': match_ratio,
                            'suggestion': f"Expand your coverage of {topic}. Currently at {int(match_ratio * 100)}% coverage."
                        })
                    # Missing topics (less than 40% match) need to be added
                    else:
                        suggestions['missing_topics'].append({
                            'topic': topic,
                            'description': description
                        })
                else:
                    # Fallback if topic not found in coverage data
                    suggestions['missing_topics'].append({
                        'topic': topic,
                        'description': description
                    })
        
        # Check for unanswered questions
        for question in term_data.get('questions', []):
            core_question = question.replace('?', '').lower()
            question_words = re.findall(r'\b\w+\b', core_question)
            significant_words = [w for w in question_words if len(w) > 3 and w not in ['what', 'when', 'where', 'which', 'who', 'why', 'how']]
            
            matches = sum(1 for word in significant_words if word in content.lower())
            
            if not significant_words or matches < len(significant_words) * 0.7:
                suggestions['unanswered_questions'].append(question)
        
        # Readability suggestions
        word_count = score_data.get('details', {}).get('word_count', 0)
        
        if word_count < 300:
            suggestions['readability_suggestions'].append("Content is too short. Aim for at least 800-1200 words for most topics.")
        elif word_count < 800:
            suggestions['readability_suggestions'].append("Content may be too brief. Consider expanding to 1000+ words for better topic coverage.")
        
        # Structure suggestions based on parsing the content
        soup = BeautifulSoup(content, 'html.parser')
        headings = soup.find_all(['h1', 'h2', 'h3', 'h4', 'h5', 'h6'])
        paragraphs = soup.find_all('p')
        
        if len(headings) < 3:
            suggestions['structure_suggestions'].append("Add more section headings to improve structure and readability.")
        
        if paragraphs:
            # Check paragraph length
            long_paragraphs = sum(1 for p in paragraphs if len(p.get_text().split()) > 200)
            if long_paragraphs > 0:
                suggestions['structure_suggestions'].append(f"Break up {long_paragraphs} long paragraph(s) into smaller chunks for better readability.")
            
            # Check for use of lists
            lists = soup.find_all(['ul', 'ol'])
            if len(lists) == 0:
                suggestions['structure_suggestions'].append("Consider adding bulleted or numbered lists to improve scannability.")
        
        return suggestions, True
    
    except Exception as e:
        error_msg = f"Exception in get_content_improvement_suggestions: {str(e)}"
        logger.error(error_msg)
        return {}, False

def create_content_scoring_brief(keyword: str, term_data: Dict, score_data: Dict, suggestions: Dict) -> BytesIO:
    """
    Create a downloadable content scoring brief with recommendations
    Returns: document_stream
    """
    try:
        doc = Document()
        
        # Add document title
        doc.add_heading(f'Content Optimization Brief: {keyword}', 0)
        
        # Add date
        doc.add_paragraph(f"Generated on: {time.strftime('%Y-%m-%d %H:%M:%S')}")
        
        # Overall Score
        doc.add_heading('Content Score', level=1)
        score_para = doc.add_paragraph()
        score_para.add_run(f"Overall Score: ").bold = True
        score_run = score_para.add_run(f"{score_data.get('overall_score', 0)} ({score_data.get('grade', 'F')})")
        
        overall_score = score_data.get('overall_score', 0)
        if overall_score >= 70:
            score_run.font.color.rgb = RGBColor(0, 128, 0)  # Green
        elif overall_score < 50:
            score_run.font.color.rgb = RGBColor(255, 0, 0)  # Red
        else:
            score_run.font.color.rgb = RGBColor(255, 165, 0)  # Orange

        # Component scores
        components = score_data.get('components', {})
        table = doc.add_table(rows=1, cols=2)
        table.style = 'Table Grid'
        
        header_cells = table.rows[0].cells
        header_cells[0].text = 'Component'
        header_cells[1].text = 'Score'
        
        for component, score in components.items():
            formatted_component = component.replace('_score', '').replace('_', ' ').title()
            row_cells = table.add_row().cells
            row_cells[0].text = formatted_component
            row_cells[1].text = str(score)
        
        # Primary Terms to Include
        doc.add_heading('Primary Terms to Include', level=1)
        
        primary_terms_table = doc.add_table(rows=1, cols=4)
        primary_terms_table.style = 'Table Grid'
        
        header_cells = primary_terms_table.rows[0].cells
        header_cells[0].text = 'Term'
        header_cells[1].text = 'Importance'
        header_cells[2].text = 'Recommended Usage'
        header_cells[3].text = 'Current Usage'
        
        primary_term_counts = score_data.get('details', {}).get('primary_term_counts', {})
        
        for term_info in term_data.get('primary_terms', []):
            term = term_info.get('term', '')
            importance = term_info.get('importance', 0)
            recommended = term_info.get('recommended_usage', 1)
            
            current_count = 0
            if term in primary_term_counts:
                current_count = primary_term_counts[term].get('count', 0)
            
            row_cells = primary_terms_table.add_row().cells
            row_cells[0].text = term
            row_cells[1].text = f"{importance:.2f}"
            row_cells[2].text = str(recommended)
            row_cells[3].text = str(current_count)
            
            # Highlight issues
            if current_count == 0:
                for cell in row_cells:
                    for paragraph in cell.paragraphs:
                        for run in paragraph.runs:
                            run.font.color.rgb = RGBColor(255, 0, 0)  # Red
            elif current_count < recommended:
                for cell in row_cells:
                    for paragraph in cell.paragraphs:
                        for run in paragraph.runs:
                            run.font.color.rgb = RGBColor(255, 165, 0)  # Orange
        
        # Content Gaps
        doc.add_heading('Content Gaps to Address', level=1)
        
        # Missing Topics
        if suggestions.get('missing_topics'):
            doc.add_heading('Missing Topics', level=2)
            for topic in suggestions.get('missing_topics', []):
                topic_para = doc.add_paragraph(style='List Bullet')
                topic_para.add_run(topic.get('topic', '')).bold = True
                topic_para.add_run(f": {topic.get('description', '')}")
        
        # Unanswered Questions
        if suggestions.get('unanswered_questions'):
            doc.add_heading('Questions to Answer', level=2)
            for question in suggestions.get('unanswered_questions', []):
                q_para = doc.add_paragraph(style='List Bullet')
                q_para.add_run(question)
        
        # Structure Recommendations
        if suggestions.get('structure_suggestions') or suggestions.get('readability_suggestions'):
            doc.add_heading('Structure & Readability Recommendations', level=1)
            
            for suggestion in suggestions.get('structure_suggestions', []):
                s_para = doc.add_paragraph(style='List Bullet')
                s_para.add_run(suggestion)
            
            for suggestion in suggestions.get('readability_suggestions', []):
                r_para = doc.add_paragraph(style='List Bullet')
                r_para.add_run(suggestion)
        
        # Save document to memory stream
        doc_stream = BytesIO()
        doc.save(doc_stream)
        doc_stream.seek(0)
        
        return doc_stream
    
    except Exception as e:
        error_msg = f"Exception in create_content_scoring_brief: {str(e)}"
        logger.error(error_msg)
        return BytesIO()

###############################################################################
# 6. Meta Title and Description Generation
###############################################################################

def generate_meta_tags(keyword: str, semantic_structure: Dict, related_keywords: List[Dict], term_data: Dict, 
                      anthropic_api_key: str) -> Tuple[str, str, bool]:
    """
    Generate optimized meta title and description for the content
    Returns: meta_title, meta_description, success_status
    """
    try:
        client = anthropic.Anthropic(api_key=anthropic_api_key)
        
        # Extract H1 and first few sections for context
        h1 = semantic_structure.get('h1', f"Complete Guide to {keyword}")
        
        # Get top 5 related keywords
        top_keywords = ", ".join([kw.get('keyword', '') for kw in related_keywords[:5] if kw.get('keyword')])
        
        # Get primary terms if available
        primary_terms = []
        if term_data and 'primary_terms' in term_data:
            primary_terms = [term.get('term') for term in term_data.get('primary_terms', [])[:5]]
        
        primary_terms_str = ", ".join(primary_terms) if primary_terms else top_keywords
        
        # Generate meta tags
        response = client.messages.create(
            model="claude-3-7-sonnet-20250219",
            max_tokens=300,
            system="You are an SEO specialist who creates optimized meta tags.",
            messages=[
                {"role": "user", "content": f"""
                Create an SEO-optimized meta title and description for an article about "{keyword}".
                
                The article's main heading is: "{h1}"
                
                Primary terms to include: {primary_terms_str}
                Related keywords to consider: {top_keywords}
                
                Guidelines:
                1. Meta title: 50-60 characters, include primary keyword near the beginning
                2. Meta description: 150-160 characters, include primary and secondary keywords
                3. Be compelling, accurate, and include a call to action in the description
                4. Avoid clickbait, use natural language
                
                Format your response as JSON:
                {{
                    "meta_title": "Your optimized meta title here",
                    "meta_description": "Your optimized meta description here"
                }}
                """}
            ],
            temperature=0.7
        )
        
        # Extract and parse JSON response
        content = response.content[0].text
        # Find JSON content within response (in case there's additional text)
        json_match = re.search(r'({.*})', content, re.DOTALL)
        if json_match:
            content = json_match.group(1)
        
        meta_data = json.loads(content)
        
        # Extract meta tags
        meta_title = meta_data.get('meta_title', f"{h1} | Your Ultimate Guide")
        meta_description = meta_data.get('meta_description', f"Learn everything about {keyword} in our comprehensive guide. Discover tips, best practices, and expert advice to master {keyword} today.")
        
        # Truncate if too long
        if len(meta_title) > 60:
            meta_title = meta_title[:57] + "..."
        
        if len(meta_description) > 160:
            meta_description = meta_description[:157] + "..."
        
        return meta_title, meta_description, True
    
    except Exception as e:
        error_msg = f"Exception in generate_meta_tags: {str(e)}"
        logger.error(error_msg)
        logger.error(traceback.format_exc())
        return f"{keyword} - Complete Guide", f"Learn everything about {keyword} in our comprehensive guide. Discover expert tips and best practices.", False

###############################################################################
# 7. Embeddings and Semantic Analysis
###############################################################################

def generate_embedding(text: str, openai_api_key: str, model: str = "text-embedding-3-small") -> Tuple[List[float], bool]:
    """
    Generate embedding for text using OpenAI API
    Returns: embedding, success_status
    """
    try:
        openai.api_key = openai_api_key
        
        # Limit text length to prevent token limit issues
        text = text[:10000]  # Adjust limit as needed
        
        response = openai.Embedding.create(
            model=model,
            input=text
        )
        
        embedding = response['data'][0]['embedding']
        return embedding, True
    
    except Exception as e:
        error_msg = f"Exception in generate_embedding: {str(e)}"
        logger.error(error_msg)
        return [], False

def analyze_semantic_structure(contents: List[Dict], anthropic_api_key: str) -> Tuple[Dict, bool]:
    """
    Analyze semantic structure of content to determine optimal hierarchy
    Returns: semantic_analysis, success_status
    """
    try:
        client = anthropic.Anthropic(api_key=anthropic_api_key)
        
        # Combine all content for analysis
        combined_content = "\n\n".join([c.get('content', '') for c in contents if c.get('content')])
        
        # Prepare summarized content if it's too long
        if len(combined_content) > 10000:
            combined_content = combined_content[:10000]
        
        # Use Claude to analyze content and suggest headings structure
        response = client.messages.create(
            model="claude-3-7-sonnet-20250219",
            max_tokens=1000,
            system="You are an SEO expert specializing in content structure.",
            messages=[
                {"role": "user", "content": f"""
                Analyze the following content from top-ranking pages and recommend an optimal semantic hierarchy 
                for a new article on this topic. Include:
                
                1. A recommended H1 title
                2. 5-7 H2 section headings
                3. 2-3 H3 subheadings under each H2
                
                Format your response as JSON:
                {{
                    "h1": "Recommended H1 Title",
                    "sections": [
                        {{
                            "h2": "First H2 Section",
                            "subsections": [
                                {{"h3": "First H3 Subsection"}},
                                {{"h3": "Second H3 Subsection"}}
                            ]
                        }},
                        ...more sections...
                    ]
                }}
                
                Content to analyze:
                {combined_content}
                """}
            ],
            temperature=0.3
        )
        
        # Extract and parse JSON response
        content = response.content[0].text
        # Find JSON content within response (in case there's additional text)
        json_match = re.search(r'({.*})', content, re.DOTALL)
        if json_match:
            content = json_match.group(1)
        
        semantic_analysis = json.loads(content)
        return semantic_analysis, True
    
    except Exception as e:
        error_msg = f"Exception in analyze_semantic_structure: {str(e)}"
        logger.error(error_msg)
        return {}, False

###############################################################################
# 8. Content Generation
###############################################################################

def generate_article(keyword: str, semantic_structure: Dict, related_keywords: List[Dict], 
                     serp_features: List[Dict], paa_questions: List[Dict], term_data: Dict, 
                     anthropic_api_key: str, guidance_only: bool = False) -> Tuple[str, bool]:
    """
    Generate comprehensive article with natural language flow and balanced keyword usage.
    If guidance_only is True, will generate writing guidance instead of full content.
    Uses Claude 3.7 Sonnet to optimize for important terms and proper length.
    Returns: article_content, success_status
    """
    try:
        client = anthropic.Anthropic(api_key=anthropic_api_key)
        
        # Ensure semantic_structure is valid
        if not semantic_structure:
            semantic_structure = {"h1": f"Guide to {keyword}", "sections": []}
        
        # Get default H1 if not present
        h1 = semantic_structure.get('h1', f"Complete Guide to {keyword}")
        
        # Use ALL sections and subsections from the semantic structure
        sections = semantic_structure.get('sections', [])
        
        # Prepare section structure with error handling
        sections_str = ""
        for section in sections:
            if section and isinstance(section, dict) and 'h2' in section:
                sections_str += f"- {section.get('h2')}\n"
                # Include all subsections
                for subsection in section.get('subsections', []):
                    if subsection and isinstance(subsection, dict) and 'h3' in subsection:
                        sections_str += f"  - {subsection.get('h3')}\n"
        
        # Add default section if none exist
        if not sections_str:
            sections_str = f"- Introduction to {keyword}\n- Key Benefits\n- How to Use\n- Conclusion\n"
        
        # Prepare related keywords with error handling
        related_kw_list = []
        if related_keywords and isinstance(related_keywords, list):
            for kw in related_keywords[:5]:  # LIMIT to top 5 only
                if kw and isinstance(kw, dict) and 'keyword' in kw:
                    related_kw_list.append(kw.get('keyword', ''))
        
        # Add default keywords if none exist
        if not related_kw_list:
            related_kw_list = [f"{keyword} guide", f"best {keyword}", f"{keyword} tips"]
        
        related_kw_str = ", ".join(related_kw_list)
        
        # Prepare SERP features with error handling
        serp_features_list = []
        if serp_features and isinstance(serp_features, list):
            for feature in serp_features[:3]:  # LIMIT to top 3 only
                if feature and isinstance(feature, dict) and 'feature_type' in feature:
                    count = feature.get('count', 1)
                    serp_features_list.append(f"{feature.get('feature_type')} ({count})")
        
        # Add default features if none exist
        if not serp_features_list:
            serp_features_list = ["featured snippet", "people also ask"]
        
        serp_features_str = ", ".join(serp_features_list)
        
        # Prepare People Also Asked questions - LIMIT to top 3
        paa_str = ""
        if paa_questions and isinstance(paa_questions, list):
            for i, question in enumerate(paa_questions[:3], 1):
                if question and isinstance(question, dict) and 'question' in question:
                    paa_str += f"{i}. {question.get('question', '')}\n"
        
        # IMPROVED: Better format for primary and secondary terms with their recommended usage
        # LIMIT to fewer terms to prevent overloading
        primary_terms_with_usage = []
        if term_data and 'primary_terms' in term_data:
            for term_info in term_data.get('primary_terms', [])[:5]:  # LIMIT to top 5 primary terms
                term = term_info.get('term', '')
                importance = term_info.get('importance', 0)
                usage = term_info.get('recommended_usage', 1)
                if term:
                    primary_terms_with_usage.append({
                        'term': term,
                        'importance': importance,
                        'usage': usage
                    })
        
        # Format primary terms for better inclusion in the prompt
        primary_terms_list = []
        for term_info in primary_terms_with_usage:
            primary_terms_list.append(f"{term_info['term']} (use {term_info['usage']} times)")
        
        primary_terms_str = "\n".join([f"- {term}" for term in primary_terms_list])
        
        # IMPROVED: Better format for secondary terms - LIMIT to fewer terms
        secondary_terms_list = []
        if term_data and 'secondary_terms' in term_data:
            for term_info in term_data.get('secondary_terms', [])[:8]:  # LIMIT to top 8 secondary terms
                term = term_info.get('term', '')
                importance = term_info.get('importance', 0)
                if term and importance > 0.5:
                    secondary_terms_list.append(term)
        
        secondary_terms_str = "\n".join([f"- {term}" for term in secondary_terms_list])
        
        if guidance_only:
            # Generate writing guidance for each section
            response = client.messages.create(
                model="claude-3-7-sonnet-20250219",
                max_tokens=2500,
                system="You are an expert SEO content strategist who provides detailed writing guidance.",
                messages=[
                    {"role": "user", "content": f"""
                    Create detailed writing guidance for an article about "{keyword}" following the semantic structure below.
                    
                    For each section (H1, H2s, and H3s), provide:
                    1. The key points to cover
                    2. Relevant statistics or data to mention (if applicable)
                    3. Tone and approach recommendations
                    4. Specific keywords to include
                    5. Approximate word count target
                    
                    Use this semantic structure:
                    H1: {h1}
                    
                    Sections:
                    {sections_str}
                    
                    Content context:
                    - Main keyword: {keyword}
                    - Related keywords to incorporate: {related_kw_str}
                    - Optimize for these SERP features: {serp_features_str}
                    - Questions to address: {paa_str}
                    
                    Important terms to include:
                    Primary terms (use these multiple times):
                    {primary_terms_str}
                    
                    Secondary terms (try to include these at least once):
                    {secondary_terms_str}
                    
                    Format the guidance with proper HTML:
                    - Main title in <h1> tags
                    - Section headings in <h2> tags
                    - Subsection headings in <h3> tags
                    - Guidance points in <p> tags
                    - Use <ul>, <li> for bullet points
                    
                    Aim for comprehensive guidance that will help a writer create a 1,200-1,500 word article.
                    """}
                ],
                temperature=0.5
            )
            
            guidance_content = response.content[0].text
            return guidance_content, True
        else:
            # SIGNIFICANTLY IMPROVED: Increased max_tokens, keep structure instructions
            response = client.messages.create(
                model="claude-3-7-sonnet-20250219",
                max_tokens=6000,  # INCREASED max_tokens to ensure full article generation
                system="""You are an expert content writer who creates concise, structured articles.
                You are known for completing articles within the specified word count while covering all requested sections.
                
                Your writing principles:
                1. Write extremely concise paragraphs (2-3 sentences each)
                2. Use clear headings for organization (H1, H2, H3, H4)
                3. Cover all requested sections briefly rather than some sections in depth
                4. Integrate required SEO terms naturally throughout the article
                5. Prioritize completeness over depth""",
                
                messages=[
                    {"role": "user", "content": f"""
                    Write a concise article about "{keyword}" that covers ALL the sections outlined below.
                    
                    Use this semantic structure:
                    H1: {h1}
                    
                    Sections to include (YOU MUST INCLUDE ALL SECTIONS LISTED - this is critical):
                    {sections_str}
                    
                    STRICT LENGTH REQUIREMENTS:
                    1. TOTAL ARTICLE LENGTH: 1,200-1,500 words maximum (STRICTLY ENFORCE THIS)
                    2. PARAGRAPH LENGTH: Each paragraph must be only 2-3 sentences (VERY IMPORTANT)
                    
                    SECTION WORD COUNTS (to ensure completeness):
                    - Introduction: 100-150 words
                    - Each H2 section: 75-100 words maximum
                    - Each H3 subsection: 50-75 words maximum
                    - Each H4 subsection: 25-50 words maximum
                    - Conclusion: 100 words maximum
                    
                    CONTENT STRUCTURE:
                    1. Start with a brief introduction
                    2. Include ALL the H2 sections listed above (crucial)
                    3. Include ALL H3 subsections listed under each H2 (crucial)
                    4. Add H4 subheadings where helpful for organization
                    5. End with a brief conclusion
                    
                    CRITICAL SEO REQUIREMENTS:
                    Primary terms to include (with exact usage count):
                    {primary_terms_str}
                    
                    Secondary terms to include (at least once each):
                    {secondary_terms_str}
                    
                    Address these questions briefly within the content:
                    {paa_str}
                    
                    CRITICAL WRITING INSTRUCTIONS:
                    1. DO NOT use rhetorical questions
                    2. ENSURE every paragraph is only 2-3 sentences
                    3. ENSURE all listed sections are included
                    4. MAINTAIN STRICT TOTAL WORD COUNT of 1,200-1,500 words
                    5. USE bullet points instead of long paragraphs for lists
                    
                    Format the article with proper HTML:
                    - Main title in <h1> tags
                    - Section headings in <h2> tags
                    - Subsection headings in <h3> tags
                    - Sub-subsection headings in <h4> tags
                    - Paragraphs in <p> tags
                    - Use <ul>, <li> for bullet points and <ol>, <li> for numbered lists
                    """}
                ],
                temperature=0.4  # Reduced temperature for more consistent output
            )
            
            article_content = response.content[0].text
            return article_content, True
    
    except Exception as e:
        error_msg = f"Exception in generate_article: {str(e)}"
        logger.error(error_msg)
        return "", False

###############################################################################
# 9. Internal Linking
###############################################################################

def parse_site_pages_spreadsheet(uploaded_file) -> Tuple[List[Dict], bool]:
    """
    Parse uploaded CSV/Excel with site pages
    Returns: pages, success_status
    """
    try:
        # Determine file type and read accordingly
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        elif uploaded_file.name.endswith(('.xlsx', '.xls')):
            df = pd.read_excel(uploaded_file)
        else:
            logger.error(f"Unsupported file type: {uploaded_file.name}")
            return [], False
        
        # Check required columns
        required_columns = ['URL', 'Title', 'Meta Description']
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            logger.error(f"Missing required columns: {', '.join(missing_columns)}")
            return [], False
        
        # Convert dataframe to list of dicts
        pages = []
        for _, row in df.iterrows():
            pages.append({
                'url': row['URL'],
                'title': row['Title'],
                'description': row['Meta Description']
            })
        
        return pages, True
    
    except Exception as e:
        error_msg = f"Exception in parse_site_pages_spreadsheet: {str(e)}"
        logger.error(error_msg)
        return [], False

def embed_site_pages(pages: List[Dict], openai_api_key: str, batch_size: int = 10) -> Tuple[List[Dict], bool]:
    """
    Generate embeddings for site pages in batches for faster processing
    Returns: pages_with_embeddings, success_status
    """
    try:
        openai.api_key = openai_api_key
        
        # Prepare texts to embed
        texts = []
        for page in pages:
            # Combine URL, title and description for embedding
            combined_text = f"{page['url']} {page['title']} {page['description']}"
            texts.append(combined_text)
        
        # Process in batches
        embeddings = []
        total_batches = (len(texts) + batch_size - 1) // batch_size
        
        for i in range(total_batches):
            start_idx = i * batch_size
            end_idx = min(start_idx + batch_size, len(texts))
            
            batch_texts = texts[start_idx:end_idx]
            
            response = openai.Embedding.create(
                model="text-embedding-3-small",
                input=batch_texts
            )
            
            batch_embeddings = [item['embedding'] for item in response['data']]
            embeddings.extend(batch_embeddings)
        
        # Add embeddings to pages
        pages_with_embeddings = []
        for i, page in enumerate(pages):
            page_with_embedding = page.copy()
            page_with_embedding['embedding'] = embeddings[i]
            pages_with_embeddings.append(page_with_embedding)
        
        return pages_with_embeddings, True
    
    except Exception as e:
        error_msg = f"Exception in embed_site_pages: {str(e)}"
        logger.error(error_msg)
        return pages, False

def verify_semantic_match(anchor_text: str, page_title: str) -> float:
    """
    Verify and score the semantic match between anchor text and page title
    Returns a similarity score (0-1)
    """
    # Define common stop words
    stop_words = {'a', 'an', 'the', 'and', 'or', 'but', 'in', 'on', 'at', 'to', 'for', 'with', 
                  'by', 'about', 'as', 'is', 'are', 'was', 'were', 'of', 'from', 'into', 'during',
                  'after', 'before', 'above', 'below', 'between', 'under', 'over', 'through'}
    
    # Convert to lowercase and tokenize
    anchor_words = {word.lower() for word in re.findall(r'\b\w+\b', anchor_text)}
    title_words = {word.lower() for word in re.findall(r'\b\w+\b', page_title)}
    
    # Remove stop words
    anchor_meaningful = anchor_words - stop_words
    title_meaningful = title_words - stop_words
    
    if not anchor_meaningful or not title_meaningful:
        return 0.0
    
    # Find overlapping words
    overlaps = anchor_meaningful.intersection(title_meaningful)
    
    # Calculate similarity score based on overlap percentage
    # (weighted toward the anchor text's coverage of title words)
    if len(overlaps) == 0:
        return 0.0
    
    # Calculate percentage of title words covered by anchor text
    title_coverage = len(overlaps) / len(title_meaningful)
    
    # Calculate percentage of anchor text words that appear in title
    anchor_precision = len(overlaps) / len(anchor_meaningful)
    
    # Combined score (weighted toward title coverage)
    similarity = (title_coverage * 0.7) + (anchor_precision * 0.3)
    
    return similarity

def generate_internal_links_with_embeddings(article_content: str, pages_with_embeddings: List[Dict], 
                                           openai_api_key: str, word_count: int) -> Tuple[str, List[Dict], bool]:
    """
    Generate internal links using paragraph-level semantic matching with embeddings
    Returns: article_with_links, links_added, success_status
    """
    try:
        openai.api_key = openai_api_key
        
        # Calculate max links based on word count
        max_links = min(10, max(2, int(word_count / 1000) * 8))  # Ensure at least 2 links
        
        # Log for debugging
        logger.info(f"Generating up to {max_links} internal links for content with {word_count} words")
        
        # 1. Extract paragraphs from the article
        soup = BeautifulSoup(article_content, 'html.parser')
        paragraphs = []
        for p_tag in soup.find_all('p'):
            para_text = p_tag.get_text()
            if len(para_text.split()) > 15:  # Only consider paragraphs with enough content
                paragraphs.append({
                    'text': para_text,
                    'html': str(p_tag),
                    'element': p_tag
                })
        
        if not paragraphs:
            logger.warning("No paragraphs found in the article")
            return article_content, [], False
        
        # 2. Generate embeddings for each paragraph
        logger.info(f"Generating embeddings for {len(paragraphs)} paragraphs")
        paragraph_texts = [p['text'] for p in paragraphs]
        
        try:
            # Get paragraph embeddings
            response = openai.Embedding.create(
                model="text-embedding-3-small",
                input=paragraph_texts
            )
            paragraph_embeddings = [item['embedding'] for item in response['data']]
            
            # Add embeddings to paragraphs
            for i, embedding in enumerate(paragraph_embeddings):
                paragraphs[i]['embedding'] = embedding
                
        except Exception as e:
            logger.error(f"Error generating paragraph embeddings: {e}")
            return article_content, [], False
        
        # 3. Find the best page match for each paragraph
        links_to_add = []
        used_paragraphs = set()  # Track paragraphs that already have links
        used_pages = set()       # Track pages that are already linked to
        
        # Only process pages that have embeddings
        valid_pages = [p for p in pages_with_embeddings if p.get('embedding')]
        
        # For each paragraph, find the best matching page
        for para_idx, paragraph in enumerate(paragraphs):
            if len(links_to_add) >= max_links or para_idx in used_paragraphs:
                continue
                
            para_embedding = paragraph.get('embedding', [])
            if not para_embedding:
                continue
                
            # Find best matching page for this paragraph
            best_score = 0.65  # Minimum threshold for a good match
            best_page = None
            
            for page in valid_pages:
                if page['url'] in used_pages:
                    continue
                    
                page_embedding = page.get('embedding', [])
                if not page_embedding:
                    continue
                    
                # Calculate cosine similarity
                similarity = np.dot(para_embedding, page_embedding) / (
                    np.linalg.norm(para_embedding) * np.linalg.norm(page_embedding)
                )
                
                if similarity > best_score:
                    best_score = similarity
                    best_page = page
            
            # If we found a good page match
            if best_page:
                page_title = best_page.get('title', '')
                para_text = paragraph['text']
                
                # Ask Claude to identify a good anchor text from the paragraph that relates to the page title
                try:
                    client = anthropic.Anthropic(api_key=openai_api_key)
                    anchor_response = client.messages.create(
                        model="claude-3-7-sonnet-20250219",
                        max_tokens=50,
                        system="You are an expert at identifying semantically relevant anchor text for links.",
                        messages=[
                            {"role": "user", "content": f"""
                            Find the BEST 2-6 word phrase in this paragraph that would make a semantically relevant anchor text for a page titled "{page_title}".
                            
                            The anchor text MUST:
                            1. Be an EXACT substring in the paragraph (case-sensitive)
                            2. Contain at least one meaningful keyword from the page title
                            3. Make sense as clickable text
                            
                            Paragraph:
                            {para_text}
                            
                            Page title:
                            {page_title}
                            
                            Return ONLY the exact anchor text phrase, nothing else.
                            """}
                        ],
                        temperature=0.3
                    )
                    
                    anchor_text = anchor_response.content[0].text.strip()
                    anchor_text = anchor_text.strip('"\'')  # Remove quotes if present
                    
                    # Verify the anchor text exists in the paragraph
                    if anchor_text in para_text:
                        # Add to our links list
                        links_to_add.append({
                            'url': best_page['url'],
                            'anchor_text': anchor_text,
                            'paragraph_index': para_idx,
                            'similarity_score': best_score,
                            'page_title': page_title
                        })
                        
                        # Mark as used
                        used_paragraphs.add(para_idx)
                        used_pages.add(best_page['url'])
                        
                        logger.info(f"Found match: '{anchor_text}' in paragraph {para_idx} for page '{page_title}'")
                except Exception as e:
                    logger.error(f"Error identifying anchor text: {e}")
        
        # 4. Apply the links to the article
        if not links_to_add:
            logger.warning("No suitable links found to add")
            return article_content, [], False
        
        # Create a deep copy of the soup to modify
        modified_soup = BeautifulSoup(article_content, 'html.parser')
        modified_paragraphs = modified_soup.find_all('p')
        
        # Apply links
        for link in links_to_add:
            para_idx = link['paragraph_index']
            if para_idx < len(modified_paragraphs):
                p_tag = modified_paragraphs[para_idx]
                p_html = str(p_tag)
                anchor_text = link['anchor_text']
                url = link['url']
                
                # Replace the text with a linked version
                new_html = p_tag.decode_contents().replace(
                    anchor_text, 
                    f'<a href="{url}">{anchor_text}</a>', 
                    1
                )
                p_tag.clear()
                p_tag.append(BeautifulSoup(new_html, 'html.parser'))
                
                # Add context to the link info
                para_text = paragraphs[para_idx]['text']
                start_pos = max(0, para_text.find(anchor_text) - 30)
                end_pos = min(len(para_text), para_text.find(anchor_text) + len(anchor_text) + 30)
                context = "..." + para_text[start_pos:end_pos].replace(anchor_text, f"[{anchor_text}]") + "..."
                
                # Update the link details
                link['context'] = context
        
        # Format for return
        links_output = []
        for link in links_to_add:
            links_output.append({
                "url": link['url'],
                "anchor_text": link['anchor_text'],
                "context": link['context'],
                "page_title": link['page_title'],
                "similarity_score": round(link['similarity_score'], 2)
            })
        
        return str(modified_soup), links_output, True
        
    except Exception as e:
        error_msg = f"Exception in generate_internal_links_with_embeddings: {str(e)}"
        logger.error(error_msg)
        logger.error(traceback.format_exc())
        return article_content, [], False

###############################################################################
# 10. Document Generation
###############################################################################

def create_word_document(keyword: str, serp_results: List[Dict], related_keywords: List[Dict],
                        semantic_structure: Dict, article_content: str, meta_title: str, 
                        meta_description: str, paa_questions: List[Dict], term_data: Dict = None,
                        score_data: Dict = None, internal_links: List[Dict] = None, 
                        guidance_only: bool = False) -> Tuple[BytesIO, bool]:
    """
    Create Word document with all components including content score if available
    Returns: document_stream, success_status
    """
    try:
        doc = Document()
        
        # Add document title
        doc.add_heading(f'SEO Brief: {keyword}', 0)
        
        # Add date
        doc.add_paragraph(f"Generated on: {time.strftime('%Y-%m-%d %H:%M:%S')}")
        
        # Add meta title and description
        doc.add_heading('Meta Tags', level=1)
        meta_paragraph = doc.add_paragraph()
        meta_paragraph.add_run("Meta Title: ").bold = True
        meta_paragraph.add_run(meta_title)
        
        desc_paragraph = doc.add_paragraph()
        desc_paragraph.add_run("Meta Description: ").bold = True
        desc_paragraph.add_run(meta_description)
        
        # Section 1: SERP Analysis
        doc.add_heading('SERP Analysis', level=1)
        doc.add_paragraph('Top 10 Organic Results:')
        
        table = doc.add_table(rows=1, cols=4)
        table.style = 'Table Grid'
        
        # Add header row
        header_cells = table.rows[0].cells
        header_cells[0].text = 'Rank'
        header_cells[1].text = 'Title'
        header_cells[2].text = 'URL'
        header_cells[3].text = 'Page Type'
        
        # Add data rows
        for result in serp_results:
            row_cells = table.add_row().cells
            row_cells[0].text = str(result.get('rank_group', ''))
            row_cells[1].text = result.get('title', '')
            row_cells[2].text = result.get('url', '')
            row_cells[3].text = result.get('page_type', '')
        
        # Add People Also Asked questions
        if paa_questions:
            doc.add_heading('People Also Asked', level=2)
            for i, question in enumerate(paa_questions, 1):
                q_paragraph = doc.add_paragraph(style='List Number')
                q_paragraph.add_run(question.get('question', '')).bold = True
                
                # Add expanded answers if available
                for expanded in question.get('expanded', []):
                    if expanded.get('description'):
                        doc.add_paragraph(expanded.get('description', ''), style='List Bullet')
        
        # Section 2: Related Keywords
        doc.add_heading('Related Keywords', level=1)
        
        kw_table = doc.add_table(rows=1, cols=3)
        kw_table.style = 'Table Grid'
        
        # Add header row
        kw_header_cells = kw_table.rows[0].cells
        kw_header_cells[0].text = 'Keyword'
        kw_header_cells[1].text = 'Search Volume'
        kw_header_cells[2].text = 'CPC ($)'
        
        # Add data rows with safe handling of values
        for kw in related_keywords:
            row_cells = kw_table.add_row().cells
            row_cells[0].text = kw.get('keyword', '')
            
            # Safe handling of search volume
            search_volume = kw.get('search_volume')
            row_cells[1].text = str(search_volume if search_volume is not None else 0)
            
            # Safe handling of CPC
            cpc_value = kw.get('cpc')
            if cpc_value is None:
                row_cells[2].text = "$0.00"
            else:
                try:
                    row_cells[2].text = f"${float(cpc_value):.2f}"
                except (ValueError, TypeError):
                    # Handle case where CPC might be a string or other non-numeric value
                    row_cells[2].text = "$0.00"
        
        # Section 3: Important Terms (if available)
        if term_data:
            doc.add_heading('Important Terms to Include', level=1)
            
            # Primary Terms
            doc.add_heading('Primary Terms', level=2)
            primary_table = doc.add_table(rows=1, cols=3)
            primary_table.style = 'Table Grid'
            
            header_cells = primary_table.rows[0].cells
            header_cells[0].text = 'Term'
            header_cells[1].text = 'Importance'
            header_cells[2].text = 'Recommended Usage'
            
            for term in term_data.get('primary_terms', []):
                row_cells = primary_table.add_row().cells
                row_cells[0].text = term.get('term', '')
                row_cells[1].text = f"{term.get('importance', 0):.2f}"
                row_cells[2].text = str(term.get('recommended_usage', 1))
            
            # Secondary Terms
            doc.add_heading('Secondary Terms', level=2)
            secondary_table = doc.add_table(rows=1, cols=2)
            secondary_table.style = 'Table Grid'
            
            header_cells = secondary_table.rows[0].cells
            header_cells[0].text = 'Term'
            header_cells[1].text = 'Importance'
            
            for term in term_data.get('secondary_terms', [])[:15]:  # Limit to top 15
                row_cells = secondary_table.add_row().cells
                row_cells[0].text = term.get('term', '')
                row_cells[1].text = f"{term.get('importance', 0):.2f}"
        
        # Section 4: Content Score (if available)
        if score_data:
            doc.add_heading('Content Score', level=1)
            
            score_paragraph = doc.add_paragraph()
            score_paragraph.add_run(f"Overall Score: ").bold = True
            score_run = score_paragraph.add_run(f"{score_data.get('overall_score', 0)} - {score_data.get('grade', 'F')}")
            
            # Color the score based on value
            overall_score = score_data.get('overall_score', 0)
            if overall_score >= 70:
                score_run.font.color.rgb = RGBColor(0, 128, 0)  # Green
            elif overall_score < 50:
                score_run.font.color.rgb = RGBColor(255, 0, 0)  # Red
            else:
                score_run.font.color.rgb = RGBColor(255, 165, 0)  # Orange
            
            # Component scores
            doc.add_heading('Score Components', level=2)
            components = score_data.get('components', {})
            
            for component, value in components.items():
                component_para = doc.add_paragraph(style='List Bullet')
                component_name = component.replace('_score', '').replace('_', ' ').title()
                component_para.add_run(f"{component_name}: ").bold = True
                component_para.add_run(f"{value}")
        
        # Section 5: Semantic Structure
        doc.add_heading('Recommended Content Structure', level=1)
        
        doc.add_paragraph(f"Recommended H1: {semantic_structure.get('h1', '')}")
        
        for i, section in enumerate(semantic_structure.get('sections', []), 1):
            doc.add_paragraph(f"H2 Section {i}: {section.get('h2', '')}")
            
            for j, subsection in enumerate(section.get('subsections', []), 1):
                doc.add_paragraph(f"    H3 Subsection {j}: {subsection.get('h3', '')}")
        
        # Section 6: Generated Article or Guidance
        doc.add_heading('Generated Article Content', level=1)
        
        # COMPLETELY REVISED: A simpler, more direct approach to HTML parsing
        if article_content and isinstance(article_content, str):
            # Log the first part of the article content to see what we're dealing with
            logger.info(f"Article content first 500 chars: {article_content[:500]}")
            
            # Check if content contains HTML
            if '<' in article_content and '>' in article_content:
                # Parse HTML content with Beautiful Soup
                soup = BeautifulSoup(article_content, 'html.parser')
                
                # Log the HTML structure for debugging
                logger.info(f"HTML structure: {soup.prettify()[:500]}")
                
                # First extract all headings directly
                headings = {}
                for tag_name in ['h1', 'h2', 'h3', 'h4', 'h5', 'h6']:
                    level = int(tag_name[1])
                    for heading in soup.find_all(tag_name):
                        text = heading.get_text().strip()
                        if text:
                            # Add to document with correct level
                            heading_obj = doc.add_heading(text, level=level)
                            # Store in dictionary for debugging
                            if tag_name not in headings:
                                headings[tag_name] = []
                            headings[tag_name].append(text)
                            
                            # For debugging
                            logger.info(f"Added heading: {tag_name} - {text}")
                            
                            # Look for following paragraph content
                            next_sibling = heading.find_next_sibling()
                            while next_sibling and next_sibling.name not in ['h1', 'h2', 'h3', 'h4', 'h5', 'h6']:
                                if next_sibling.name == 'p':
                                    p_text = next_sibling.get_text().strip()
                                    if p_text:
                                        doc.add_paragraph(p_text)
                                elif next_sibling.name == 'ul':
                                    for li in next_sibling.find_all('li'):
                                        li_text = li.get_text().strip()
                                        if li_text:
                                            doc.add_paragraph(li_text, style='List Bullet')
                                elif next_sibling.name == 'ol':
                                    for li in next_sibling.find_all('li'):
                                        li_text = li.get_text().strip()
                                        if li_text:
                                            doc.add_paragraph(li_text, style='List Number')
                                
                                next_sibling = next_sibling.find_next_sibling()
                
                # Log the extracted headings for debugging
                for tag_name, texts in headings.items():
                    logger.info(f"Extracted {len(texts)} {tag_name} headings: {texts}")
                
                # If no headings were found, we need to try a different approach
                if not any(headings.values()):
                    logger.warning("No headings found in HTML. Trying alternative approach")
                    
                    # Alternative approach: try to split the content by newlines and look for heading patterns
                    lines = article_content.split('\n')
                    for line in lines:
                        line = line.strip()
                        if not line:
                            continue
                        
                        # Try to detect if this is a heading based on HTML tags
                        heading_match = re.search(r'<h([1-6])>(.*?)</h\1>', line)
                        if heading_match:
                            level = int(heading_match.group(1))
                            text = heading_match.group(2).strip()
                            if text:
                                doc.add_heading(text, level=level)
                                logger.info(f"Added heading via regex: h{level} - {text}")
                        # Otherwise, it's probably a paragraph
                        elif not line.startswith('<') and not line.endswith('>'):
                            # Clean up any HTML tags
                            clean_line = re.sub(r'<.*?>', '', line).strip()
                            if clean_line:
                                doc.add_paragraph(clean_line)
            else:
                # Content is plain text - try to identify headings by format
                logger.info("Content appears to be plain text, looking for heading patterns")
                lines = article_content.split('\n')
                in_list = False
                
                for i, line in enumerate(lines):
                    line = line.strip()
                    if not line:
                        continue
                    
                    # Check for Markdown-style headings
                    if line.startswith('# '):
                        doc.add_heading(line[2:], level=1)
                    elif line.startswith('## '):
                        doc.add_heading(line[3:], level=2)
                    elif line.startswith('### '):
                        doc.add_heading(line[4:], level=3)
                    elif line.startswith('#### '):
                        doc.add_heading(line[5:], level=4)
                    # Check for list items
                    elif line.startswith('- ') or line.startswith('* '):
                        doc.add_paragraph(line[2:], style='List Bullet')
                        in_list = True
                    elif line.startswith('1. ') or line.startswith('1) '):
                        doc.add_paragraph(line[3:], style='List Number')
                        in_list = True
                    # Check for potential heading (all caps, short line)
                    elif line.isupper() and len(line) < 80:
                        doc.add_heading(line, level=2)
                    # Regular paragraph
                    else:
                        # Check if this could be a non-markdown heading
                        next_line = lines[i+1].strip() if i+1 < len(lines) else ""
                        if (len(line) < 80 and next_line and 
                            (all(c == '=' for c in next_line) or all(c == '-' for c in next_line))):
                            # This looks like a heading with underline
                            level = 1 if '=' in next_line else 2
                            doc.add_heading(line, level=level)
                        else:
                            # Regular paragraph
                            doc.add_paragraph(line)
                            in_list = False
        
        # Section 7: Internal Linking (if provided)
        if internal_links:
            doc.add_heading('Internal Linking Summary', level=1)
            
            link_table = doc.add_table(rows=1, cols=3)
            link_table.style = 'Table Grid'
            
            # Add header row
            link_header_cells = link_table.rows[0].cells
            link_header_cells[0].text = 'URL'
            link_header_cells[1].text = 'Anchor Text'
            link_header_cells[2].text = 'Context'
            
            # Add data rows
            for link in internal_links:
                row_cells = link_table.add_row().cells
                row_cells[0].text = link.get('url', '')
                row_cells[1].text = link.get('anchor_text', '')
                row_cells[2].text = link.get('context', '')
        
        # Save document to memory stream
        doc_stream = BytesIO()
        doc.save(doc_stream)
        doc_stream.seek(0)
        
        return doc_stream, True
    
    except Exception as e:
        error_msg = f"Exception in create_word_document: {str(e)}"
        logger.error(error_msg)
        logger.error(traceback.format_exc())
        return BytesIO(), False

###############################################################################
# 11. Content Update Functions
###############################################################################

def parse_word_document(uploaded_file) -> Tuple[Dict, bool]:
    """
    Parse uploaded Word document to extract content structure
    Returns: document_content, success_status
    """
    try:
        # Read the document
        doc = Document(BytesIO(uploaded_file.getvalue()))
        
        # Extract content structure
        document_content = {
            'title': '',
            'headings': [],
            'paragraphs': [],
            'full_text': ''
        }
        
        # Extract text and maintain hierarchy
        full_text = []
        current_heading = None
        
        for para in doc.paragraphs:
            text = para.text.strip()
            if not text:
                continue
                
            # Check if it's a heading
            if para.style.name.startswith('Heading'):
                heading_level = int(para.style.name.replace('Heading', '')) if para.style.name != 'Heading' else 1
                current_heading = {
                    'text': text,
                    'level': heading_level,
                    'paragraphs': []
                }
                document_content['headings'].append(current_heading)
                
                # If it's the title (first heading), save it
                if heading_level == 1 and not document_content['title']:
                    document_content['title'] = text
            else:
                # It's a paragraph
                para_obj = {
                    'text': text,
                    'heading': current_heading['text'] if current_heading else None
                }
                document_content['paragraphs'].append(para_obj)
                
                # Add to current heading if available
                if current_heading:
                    current_heading['paragraphs'].append(text)
                
            full_text.append(text)
        
        document_content['full_text'] = '\n\n'.join(full_text)
        
        return document_content, True
    except Exception as e:
        error_msg = f"Exception in parse_word_document: {str(e)}"
        logger.error(error_msg)
        logger.error(traceback.format_exc())
        return {}, False

def analyze_content_gaps(existing_content: Dict, competitor_contents: List[Dict], semantic_structure: Dict, 
                        term_data: Dict, score_data: Dict, anthropic_api_key: str, 
                        keyword: str, paa_questions: List[Dict] = None) -> Tuple[Dict, bool]:
    """
    Enhanced content gap analysis that incorporates content scoring data
    Returns: content_gaps, success_status
    """
    try:
        client = anthropic.Anthropic(api_key=anthropic_api_key)
        
        # Extract existing headings
        existing_headings = [h['text'] for h in existing_content.get('headings', [])]
        
        # Extract recommended headings from semantic structure
        recommended_headings = []
        if 'h1' in semantic_structure:
            recommended_headings.append(semantic_structure['h1'])
        
        for section in semantic_structure.get('sections', []):
            if 'h2' in section:
                recommended_headings.append(section['h2'])
                for subsection in section.get('subsections', []):
                    if 'h3' in subsection:
                        recommended_headings.append(subsection['h3'])
        
        # Combine competitor content
        competitor_text = ""
        for content in competitor_contents:
            competitor_text += content.get('content', '') + "\n\n"
        
        # Truncate if too long
        if len(competitor_text) > 12000:
            competitor_text = competitor_text[:12000]
            
        # Prepare PAA questions for the prompt
        paa_text = ""
        if paa_questions:
            paa_text = "People Also Asked Questions:\n"
            for i, q in enumerate(paa_questions, 1):
                paa_text += f"{i}. {q.get('question', '')}\n"
        
        # Prepare content scoring data
        content_score_text = ""
        if score_data:
            overall_score = score_data.get('overall_score', 0)
            grade = score_data.get('grade', 'F')
            content_score_text = f"Content Score: {overall_score} ({grade})\n\n"
            
            # Add component scores
            components = score_data.get('components', {})
            content_score_text += "Component Scores:\n"
            for component, score in components.items():
                component_name = component.replace('_score', '').replace('_', ' ').title()
                content_score_text += f"- {component_name}: {score}\n"
            
            # Add missing terms
            if 'details' in score_data:
                details = score_data.get('details', {})
                
                # Missing primary terms
                primary_term_counts = details.get('primary_term_counts', {})
                missing_primary = []
                underused_primary = []
                
                if term_data and 'primary_terms' in term_data:
                    for term_info in term_data.get('primary_terms', []):
                        term = term_info.get('term', '')
                        recommended = term_info.get('recommended_usage', 1)
                        
                        if term in primary_term_counts:
                            actual = primary_term_counts[term].get('count', 0)
                            if actual == 0:
                                missing_primary.append(term)
                            elif actual < recommended:
                                underused_primary.append(f"{term} (used {actual}/{recommended} times)")
                        else:
                            missing_primary.append(term)
                
                if missing_primary:
                    content_score_text += "\nMissing Primary Terms:\n"
                    for term in missing_primary:
                        content_score_text += f"- {term}\n"
                
                if underused_primary:
                    content_score_text += "\nUnderused Primary Terms:\n"
                    for term in underused_primary:
                        content_score_text += f"- {term}\n"
                
                # Unanswered questions from scoring
                question_coverage = details.get('question_coverage', {})
                unanswered = []
                
                for question, info in question_coverage.items():
                    if not info.get('answered', False):
                        unanswered.append(question)
                
                if unanswered:
                    content_score_text += "\nUnanswered Questions from Content Scoring:\n"
                    for q in unanswered:
                        content_score_text += f"- {q}\n"
        
        # Prepare term data
        term_data_text = ""
        if term_data:
            # Primary terms
            if 'primary_terms' in term_data:
                term_data_text += "Primary Terms (Top 10):\n"
                for term_info in term_data.get('primary_terms', [])[:10]:
                    term = term_info.get('term', '')
                    importance = term_info.get('importance', 0)
                    usage = term_info.get('recommended_usage', 1)
                    term_data_text += f"- {term} (importance: {importance:.2f}, usage: {usage})\n"
            
            # Topics to cover
            if 'topics' in term_data:
                term_data_text += "\nKey Topics to Cover:\n"
                for topic_info in term_data.get('topics', []):
                    topic = topic_info.get('topic', '')
                    description = topic_info.get('description', '')
                    term_data_text += f"- {topic}: {description}\n"
        
        # Use Claude to analyze content gaps
        response = client.messages.create(
            model="claude-3-7-sonnet-20250219",
            max_tokens=2500,
            system="You are an expert SEO content analyst specializing in identifying content gaps and semantic relevancy issues.",
            messages=[
                {"role": "user", "content": f"""
                Analyze the existing content and compare it with top-performing competitor content to identify gaps for the keyword: {keyword}
                
                Existing Content Headings:
                {json.dumps(existing_headings, indent=2)}
                
                Recommended Content Structure Based on Competitors:
                {json.dumps(recommended_headings, indent=2)}
                
                {content_score_text}
                
                {term_data_text}
                
                {paa_text}
                
                Existing Content:
                {existing_content.get('full_text', '')[:5000]}
                
                Competitor Content (Sample):
                {competitor_text[:5000]}
                
                Identify:
                1. Missing headings/sections that should be added
                2. Existing headings that should be revised/renamed
                3. Key topics/points covered by competitors but missing in the existing content
                4. Content areas that need expansion
                5. SEMANTIC RELEVANCY ISSUES: Analyze if the content is too broadly focused instead of targeting the specific keyword "{keyword}". Identify sections that need to be refocused.
                6. TERM USAGE ISSUES: Identify where important terms are missing or underused.
                7. UNANSWERED QUESTIONS: If provided, analyze which "People Also Asked" questions are not adequately addressed in the content and should be incorporated.
                
                Format your response as JSON:
                {{
                    "missing_headings": [
                        {{ 
                            "heading": "Heading Text", 
                            "level": 2, 
                            "suggested_content": "Brief description of what this section should cover",
                            "insert_after": "Name of existing heading to insert after or 'START' for beginning or 'END' for end"
                        }}
                    ],
                    "revised_headings": [
                        {{ "original": "Original Heading", "suggested": "Improved Heading", "reason": "Reason for change" }}
                    ],
                    "content_gaps": [
                        {{ "topic": "Topic Name", "details": "What's missing about this topic", "suggested_content": "Suggested content to add" }}
                    ],
                    "expansion_areas": [
                        {{ "section": "Section Name", "reason": "Why this needs expansion", "suggested_content": "Additional content to include" }}
                    ],
                    "semantic_relevancy_issues": [
                        {{ 
                            "section": "Section that's off-target", 
                            "issue": "Description of how the content is too broad or off-target", 
                            "recommendation": "How to refocus the content on the keyword '{keyword}'"
                        }}
                    ],
                    "term_usage_issues": [
                        {{
                            "term": "Missing or underused term",
                            "section": "Section where it should be added",
                            "suggestion": "How to naturally incorporate this term"
                        }}
                    ],
                    "unanswered_questions": [
                        {{
                            "question": "People Also Asked question that isn't addressed",
                            "insert_into_section": "Section where answer should be added",
                            "suggested_answer": "Brief answer to include in the content"
                        }}
                    ]
                }}
                """}
            ],
            temperature=0.4
        )
        
        # Extract and parse JSON response
        content = response.content[0].text
        # Find JSON content within response (in case there's additional text)
        json_match = re.search(r'({.*})', content, re.DOTALL)
        if json_match:
            content = json_match.group(1)
        
        content_gaps = json.loads(content)
        return content_gaps, True
    
    except Exception as e:
        error_msg = f"Exception in analyze_content_gaps: {str(e)}"
        logger.error(error_msg)
        logger.error(traceback.format_exc())
        return {}, False

def create_updated_document(existing_content: Dict, content_gaps: Dict, keyword: str, score_data: Dict = None) -> Tuple[BytesIO, bool]:
    """
    Enhanced document creation with content score information
    Returns: document_stream, success_status
    """
    try:
        doc = Document()
        
        # Add title
        doc.add_heading(f'Content Update Recommendations: {keyword}', 0)
        
        # Add date
        doc.add_paragraph(f"Generated on: {time.strftime('%Y-%m-%d %H:%M:%S')}")
        
        # Add content score if available
        if score_data:
            score_section = doc.add_heading('Content Score Assessment', 1)
            
            overall_score = score_data.get('overall_score', 0)
            grade = score_data.get('grade', 'F')
            
            score_para = doc.add_paragraph()
            score_para.add_run(f"Overall Score: ").bold = True
            score_run = score_para.add_run(f"{overall_score} ({grade})")
            
            # Color the score based on value
            if overall_score >= 70:
                score_run.font.color.rgb = RGBColor(0, 128, 0)  # Green
            elif overall_score < 50:
                score_run.font.color.rgb = RGBColor(255, 0, 0)  # Red
            else:
                score_run.font.color.rgb = RGBColor(255, 165, 0)  # Orange
            
            # Component scores
            if 'components' in score_data:
                components = score_data.get('components', {})
                
                component_table = doc.add_table(rows=1, cols=2)
                component_table.style = 'Table Grid'
                
                header_cells = component_table.rows[0].cells
                header_cells[0].text = 'Component'
                header_cells[1].text = 'Score'
                
                for component, score in components.items():
                    row_cells = component_table.add_row().cells
                    component_name = component.replace('_score', '').replace('_', ' ').title()
                    row_cells[0].text = component_name
                    row_cells[1].text = str(score)
            
            # Score improvement projection
            doc.add_paragraph()
            improvement_para = doc.add_paragraph()
            improvement_para.add_run("Projected Score After Updates: ").bold = True
            
            # Calculate projected improvement
            projected_score = min(100, overall_score + 20)  # Assume ~20 point improvement
            projected_grade = get_score_grade(projected_score)
            
            projected_run = improvement_para.add_run(f"{projected_score} ({projected_grade})")
            projected_run.font.color.rgb = RGBColor(0, 128, 0)  # Green
            
            doc.add_paragraph("Implementing the recommendations in this document is projected to significantly improve your content's search optimization score.")
        
        # Executive Summary
        doc.add_heading('Executive Summary', 1)
        summary = doc.add_paragraph()
        summary.add_run(f"This document contains recommended updates to improve your content for the target keyword '{keyword}'. ")
        summary.add_run("Based on competitor analysis and search trends, we recommend the following improvements:")
        
        # Add bullet points summarizing key recommendations
        recommendations = []
        if content_gaps.get('semantic_relevancy_issues'):
            recommendations.append("Improve semantic relevancy to better target the keyword")
        if content_gaps.get('term_usage_issues'):
            recommendations.append(f"Address {len(content_gaps.get('term_usage_issues', []))} term usage issues")
        if content_gaps.get('missing_headings'):
            recommendations.append(f"Add {len(content_gaps.get('missing_headings', []))} new sections")
        if content_gaps.get('revised_headings'):
            recommendations.append(f"Revise {len(content_gaps.get('revised_headings', []))} existing headings")
        if content_gaps.get('content_gaps'):
            recommendations.append(f"Address {len(content_gaps.get('content_gaps', []))} content gaps")
        if content_gaps.get('expansion_areas'):
            recommendations.append(f"Expand {len(content_gaps.get('expansion_areas', []))} sections")
        if content_gaps.get('unanswered_questions'):
            recommendations.append(f"Address {len(content_gaps.get('unanswered_questions', []))} 'People Also Asked' questions")
        
        for rec in recommendations:
            rec_para = doc.add_paragraph(rec, style='List Bullet')
        
        # 1. Semantic Relevancy Issues Section
        if content_gaps.get('semantic_relevancy_issues'):
            doc.add_heading('Semantic Relevancy Recommendations', 1)
            
            explanation = doc.add_paragraph()
            explanation.add_run(f"Your content appears to be focused more broadly than the target keyword '{keyword}'. ")
            explanation.add_run("The following recommendations will help align your content more closely with search intent:")
            
            for issue in content_gaps.get('semantic_relevancy_issues', []):
                section = issue.get('section', '')
                doc.add_heading(section, 2)
                
                issue_para = doc.add_paragraph()
                issue_para.add_run("Issue: ").bold = True
                issue_para.add_run(issue.get('issue', ''))
                
                rec_para = doc.add_paragraph()
                rec_para.add_run("Recommendation: ").bold = True
                rec_text = rec_para.add_run(issue.get('recommendation', ''))
                rec_text.font.color.rgb = RGBColor(255, 0, 0)  # Red
                
                doc.add_paragraph()  # Add spacing
        
        # Term Usage Issues (New Section)
        if content_gaps.get('term_usage_issues'):
            doc.add_heading('Term Usage Recommendations', 1)
            
            term_intro = doc.add_paragraph()
            term_intro.add_run("To improve your content's semantic relevance and keyword targeting, add these important terms:")
            
            term_table = doc.add_table(rows=1, cols=3)
            term_table.style = 'Table Grid'
            
            # Add header row
            header_cells = term_table.rows[0].cells
            header_cells[0].text = 'Term'
            header_cells[1].text = 'Section'
            header_cells[2].text = 'Recommendation'
            
            for issue in content_gaps.get('term_usage_issues', []):
                row_cells = term_table.add_row().cells
                term_run = row_cells[0].paragraphs[0].add_run(issue.get('term', ''))
                term_run.bold = True
                
                row_cells[1].text = issue.get('section', '')
                row_cells[2].text = issue.get('suggestion', '')
            
            doc.add_paragraph()  # Add spacing
        
        # 2. Content Structure Recommendations
        doc.add_heading('Content Structure Recommendations', 1)
        
        # a. Heading Revisions
        if content_gaps.get('revised_headings'):
            doc.add_heading('Heading Revisions', 2)
            heading_intro = doc.add_paragraph()
            heading_intro.add_run("The following heading changes will improve your content's focus and SEO performance:")
            
            table = doc.add_table(rows=1, cols=3)
            table.style = 'Table Grid'
            
            # Add header row
            header_cells = table.rows[0].cells
            header_cells[0].text = 'Current Heading'
            header_cells[1].text = 'Recommended Heading'
            header_cells[2].text = 'Rationale'
            
            for revision in content_gaps.get('revised_headings', []):
                row_cells = table.add_row().cells
                
                # Current heading
                row_cells[0].text = revision.get('original', '')
                
                # Recommended heading in red
                recommended_heading = row_cells[1].paragraphs[0].add_run(revision.get('suggested', ''))
                recommended_heading.font.color.rgb = RGBColor(255, 0, 0)
                
                # Rationale
                row_cells[2].text = revision.get('reason', '')
            
            doc.add_paragraph()  # Add spacing
            
        # b. New Sections to Add
        if content_gaps.get('missing_headings'):
            doc.add_heading('Recommended New Sections', 2)
            
            sections_intro = doc.add_paragraph()
            sections_intro.add_run("Add the following sections to make your content more comprehensive:")
            
            for heading in content_gaps.get('missing_headings', []):
                heading_level = heading.get('level', 2)
                heading_text = heading.get('heading', '')
                
                # Make H3 or H4 for better document structure
                actual_level = min(heading_level + 1, 3)
                heading_para = doc.add_heading(heading_text, level=actual_level)
                heading_para.runs[0].font.color.rgb = RGBColor(255, 0, 0)  # Red
                
                # Where to insert
                position_para = doc.add_paragraph()
                position_para.add_run("Placement: ").bold = True
                position_para.add_run(f"After '{heading.get('insert_after', 'END')}' section")
                
                # Content suggestion
                if heading.get('suggested_content'):
                    content_para = doc.add_paragraph()
                    content_para.add_run("Content to include: ").bold = True
                    content_para.add_run(heading.get('suggested_content', ''))
                
                doc.add_paragraph()  # Add spacing
        
        # 3. Content Gap Recommendations
        if content_gaps.get('content_gaps') or content_gaps.get('expansion_areas'):
            doc.add_heading('Content Gap Recommendations', 1)
            
            # a. Missing Topics
            if content_gaps.get('content_gaps'):
                doc.add_heading('Key Topics to Add', 2)
                
                topics_intro = doc.add_paragraph()
                topics_intro.add_run("The following topics are covered by competitors but missing in your content:")
                
                for gap in content_gaps.get('content_gaps', []):
                    topic_heading = doc.add_heading(gap.get('topic', ''), 3)
                    topic_heading.runs[0].font.color.rgb = RGBColor(255, 0, 0)  # Red
                    
                    issue_para = doc.add_paragraph()
                    issue_para.add_run("Gap: ").bold = True
                    issue_para.add_run(gap.get('details', ''))
                    
                    if gap.get('suggested_content'):
                        content_para = doc.add_paragraph()
                        content_para.add_run("Suggested Content: ").bold = True
                        content_para.add_run(gap.get('suggested_content', ''))
                    
                    doc.add_paragraph()  # Add spacing
            
            # b. Areas to Expand
            if content_gaps.get('expansion_areas'):
                doc.add_heading('Areas to Expand', 2)
                
                expansion_intro = doc.add_paragraph()
                expansion_intro.add_run("Enhance the following sections with additional content:")
                
                for area in content_gaps.get('expansion_areas', []):
                    area_heading = doc.add_heading(area.get('section', ''), 3)
                    
                    reason_para = doc.add_paragraph()
                    reason_para.add_run("Reason for expansion: ").bold = True
                    reason_para.add_run(area.get('reason', ''))
                    
                    if area.get('suggested_content'):
                        content_para = doc.add_paragraph()
                        content_para.add_run("Content to add: ").bold = True
                        content_para.add_run(area.get('suggested_content', ''))
                    
                    doc.add_paragraph()  # Add spacing
        
        # 4. People Also Asked Recommendations
        if content_gaps.get('unanswered_questions'):
            doc.add_heading('People Also Asked Questions to Address', 1)
            
            paa_intro = doc.add_paragraph()
            paa_intro.add_run("Incorporate answers to these common questions to improve your content's search relevance:")
            
            for question in content_gaps.get('unanswered_questions', []):
                q_para = doc.add_paragraph()
                q_para.add_run("Question: ").bold = True
                q_text = q_para.add_run(question.get('question', ''))
                q_text.italic = True
                
                if question.get('insert_into_section'):
                    section_para = doc.add_paragraph()
                    section_para.add_run("Recommended section: ").bold = True
                    section_para.add_run(question.get('insert_into_section', ''))
                
                if question.get('suggested_answer'):
                    answer_para = doc.add_paragraph()
                    answer_para.add_run("Suggested answer: ").bold = True
                    answer_para.add_run(question.get('suggested_answer', ''))
                
                doc.add_paragraph()  # Add spacing
                
        # 5. Implementation Guide
        doc.add_heading('Implementation Guide', 1)
        
        guide_para = doc.add_paragraph()
        guide_para.add_run("To implement these recommendations effectively:").bold = True
        
        implementation_steps = [
            "Start by addressing the semantic relevancy issues to align your content with the target keyword",
            "Add missing important terms to improve keyword targeting and relevance",
            "Update headings to improve clarity and search relevance",
            "Add missing sections in the recommended locations",
            "Incorporate answers to 'People Also Asked' questions to address search intent",
            "Expand thin areas with additional, valuable content",
            "After making these changes, review the content as a whole to ensure natural flow and consistency"
        ]
        
        for step in implementation_steps:
            step_para = doc.add_paragraph(step, style='List Number')
        
        # Save document to memory stream
        doc_stream = BytesIO()
        doc.save(doc_stream)
        doc_stream.seek(0)
        
        return doc_stream, True
    
    except Exception as e:
        error_msg = f"Exception in create_updated_document: {str(e)}"
        logger.error(error_msg)
        logger.error(traceback.format_exc())
        return BytesIO(), False

def generate_optimized_article_with_tracking(existing_content: Dict, competitor_contents: List[Dict], 
                              semantic_structure: Dict, related_keywords: List[Dict],
                              keyword: str, paa_questions: List[Dict], term_data: Dict,
                              anthropic_api_key: str, target_word_count: int = 1800) -> Tuple[str, str, bool]:
    """
    Enhanced article generation that incorporates term data using Claude 3.7 Sonnet
    Returns: optimized_html_content, change_summary, success_status
    """
    try:
        client = anthropic.Anthropic(api_key=anthropic_api_key)
        
        # Extract existing content structure
        original_content = existing_content.get('full_text', '')
        existing_headings = existing_content.get('headings', [])
        
        # Estimate words per section based on target word count and structure
        num_sections = len(semantic_structure.get('sections', []))
        if num_sections == 0:
            num_sections = 5  # Default if no sections defined
        
        # Calculate target words per section (allowing 15% for H1 and conclusion)
        words_per_section = int((target_word_count * 0.85) / num_sections)
        
        # Get section text for each heading to process sections individually
        section_content = {}
        for i, heading in enumerate(existing_headings):
            heading_text = heading.get('text', '')
            paragraphs = heading.get('paragraphs', [])
            section_text = "\n\n".join(paragraphs)
            section_content[heading_text] = section_text
        
        # Process the content section by section
        optimized_sections = []
        
        # Track changes by category
        structure_changes = []
        content_additions = []
        content_improvements = []
        
        # First, generate the new structure
        h1 = semantic_structure.get('h1', f"Complete Guide to {keyword}")
        optimized_sections.append(f"<h1>{h1}</h1>")
        
        # Track word count
        current_word_count = len(h1.split())
        
        # Track what original headings have been processed
        processed_headings = set()
        
        # Prepare important terms data
        primary_terms_str = ""
        if term_data and 'primary_terms' in term_data:
            primary_terms = []
            for term_info in term_data.get('primary_terms', [])[:10]:  # Top 10 primary terms
                term = term_info.get('term', '')
                importance = term_info.get('importance', 0)
                usage = term_info.get('recommended_usage', 1)
                if term:
                    primary_terms.append(f"{term} (importance: {importance:.2f}, usage: {usage})")
            
            primary_terms_str = "\n".join(primary_terms)
        
        secondary_terms_str = ""
        if term_data and 'secondary_terms' in term_data:
            secondary_terms = []
            for term_info in term_data.get('secondary_terms', [])[:15]:  # Top 15 secondary terms
                term = term_info.get('term', '')
                importance = term_info.get('importance', 0)
                if term and importance > 0.5:
                    secondary_terms.append(f"{term} (importance: {importance:.2f})")
            
            secondary_terms_str = "\n".join(secondary_terms)
        
        topics_to_cover_str = ""
        if term_data and 'topics' in term_data:
            topics = []
            for topic_info in term_data.get('topics', []):
                topic = topic_info.get('topic', '')
                description = topic_info.get('description', '')
                if topic:
                    topics.append(f"{topic}: {description}")
            
            topics_to_cover_str = "\n".join(topics)
        
        # For each recommended section in the new structure
        for section in semantic_structure.get('sections', []):
            h2 = section.get('h2', '')
            if not h2:
                continue
                
            # Skip sections if we're already approaching the target word count
            if current_word_count > (target_word_count * 0.85):
                break
                
            # Find most relevant original heading for this section
            matching_response = client.messages.create(
                model="claude-3-7-sonnet-20250219",
                max_tokens=50,
                system="You are an expert at matching content sections.",
                messages=[
                    {"role": "user", "content": f"""
                        Find the most relevant heading from the original content that matches this new section:
                        
                        New section: {h2}
                        
                        Original headings to choose from:
                        {json.dumps([h.get('text', '') for h in existing_headings if h.get('text', '') not in processed_headings], indent=2)}
                        
                        Return ONLY the exact text of the best matching original heading, or "NONE" if no good match exists.
                    """}
                ],
                temperature=0.1
            )
            
            matching_heading = matching_response.content[0].text.strip()
            optimized_sections.append(f"<h2>{h2}</h2>")
            current_word_count += len(h2.split())
            
            # Calculate words available for this section and its subsections
            subsection_count = len(section.get('subsections', []))
            # If there are subsections, allocate 60% to the main section and 40% to subsections
            if subsection_count > 0:
                section_word_limit = int(words_per_section * 0.6)
                subsection_word_limit = int((words_per_section * 0.4) / subsection_count)
            else:
                section_word_limit = words_per_section
                subsection_word_limit = 0
            
            if matching_heading == "NONE" or matching_heading not in section_content:
                # No matching content found - create new section with controlled length
                section_content_response = client.messages.create(
                    model="claude-3-7-sonnet-20250219",
                    max_tokens=section_word_limit * 2,  # Allow some extra tokens for HTML
                    system="You are an expert content writer focused on concise, informative content.",
                    messages=[
                        {"role": "user", "content": f"""
                            Write content for this section about "{keyword}": {h2}
                            
                            Requirements:
                            1. Include relevant information based on competitor content
                            2. Improve semantic relevance to the keyword
                            3. STRICTLY limit to {section_word_limit} words
                            4. Create substantive, non-fluff content
                            
                            Important terms to include:
                            Primary terms (use these if relevant):
                            {primary_terms_str}
                            
                            Secondary terms (try to include these if relevant):
                            {secondary_terms_str}
                            
                            Key topics to cover (if relevant to this section):
                            {topics_to_cover_str}
                            
                            Format with proper HTML paragraph tags.
                        """}
                    ],
                    temperature=0.4
                )
                
                new_section_content = section_content_response.content[0].text
                optimized_sections.append(new_section_content)
                
                # Estimate added word count
                section_words = len(re.findall(r'\b\w+\b', new_section_content))
                current_word_count += section_words
                
                # Track change
                content_additions.append(f"Added new section '{h2}' based on competitor analysis")
                structure_changes.append(f"Added new section: {h2}")
                
            else:
                # Found matching content - preserve and enhance
                processed_headings.add(matching_heading)
                original_section_content = section_content.get(matching_heading, '')
                
                # Enhance this section with strict word count limit
                enhanced_section_response = client.messages.create(
                    model="claude-3-7-sonnet-20250219",
                    max_tokens=section_word_limit * 2,  # Allow some extra tokens for HTML and improvements
                    system="You are an expert at enhancing content while preserving value and maintaining conciseness.",
                    messages=[
                        {"role": "user", "content": f"""
                            Enhance this original content section while PRESERVING its value.
                            
                            Original heading: {matching_heading}
                            New heading: {h2}
                            
                            Original content:
                            {original_section_content}
                            
                            Instructions:
                            1. Keep all valuable information from the original content
                            2. Preserve specific examples, data points, and unique insights
                            3. Improve semantic relevance to keyword "{keyword}"
                            4. Fix any unclear writing but maintain the original voice
                            5. STRICTLY limit to {section_word_limit} words
                            
                            Important terms to include or increase usage of:
                            Primary terms (use these if relevant):
                            {primary_terms_str}
                            
                            Secondary terms (try to include these if relevant):
                            {secondary_terms_str}
                            
                            Format with proper HTML paragraph tags.
                            
                            Also provide a single sentence summary of the key improvements you made:
                            IMPROVEMENTS: [single sentence summary of key improvements]
                        """}
                    ],
                    temperature=0.3
                )
                
                enhanced_response = enhanced_section_response.content[0].text
                
                # Extract improvements summary
                improvements_summary = ""
                if "IMPROVEMENTS:" in enhanced_response:
                    content_parts = enhanced_response.split("IMPROVEMENTS:")
                    enhanced_content = content_parts[0].strip()
                    improvements_summary = content_parts[1].strip()
                    
                    # Add to improvement tracking
                    if matching_heading != h2:
                        structure_changes.append(f"Renamed '{matching_heading}' to '{h2}'")
                    
                    content_improvements.append(f"Enhanced '{h2}': {improvements_summary}")
                else:
                    enhanced_content = enhanced_response
                    if matching_heading != h2:
                        structure_changes.append(f"Renamed '{matching_heading}' to '{h2}'")
                
                optimized_sections.append(enhanced_content)
                
                # Estimate added word count
                section_words = len(re.findall(r'\b\w+\b', enhanced_content))
                current_word_count += section_words
            
            # Process H3 subsections - but only if we have word count budget remaining
            if current_word_count < target_word_count:
                # Only process up to 3 subsections per section to control length
                for subsection in section.get('subsections', [])[:3]:
                    h3 = subsection.get('h3', '')
                    if not h3:
                        continue
                    
                    # Skip if we're getting too close to target word count
                    if current_word_count > (target_word_count * 0.95):
                        break
                    
                    optimized_sections.append(f"<h3>{h3}</h3>")
                    current_word_count += len(h3.split())
                    
                    # Generate content for this subsection with strict word limit
                    subsection_content_response = client.messages.create(
                        model="claude-3-7-sonnet-20250219",
                        max_tokens=subsection_word_limit * 2,  # Allow some extra tokens for HTML
                        system="You are an expert content writer focused on brevity and impact.",
                        messages=[
                            {"role": "user", "content": f"""
                                Write concise content for this subsection about "{keyword}": {h3} (under main section {h2})
                                
                                Requirements:
                                1. Include ONLY the most essential information
                                2. Improve semantic relevance to keyword
                                3. STRICTLY limit to {subsection_word_limit} words total
                                4. Be substantive and informative despite brevity
                                
                                Important terms to include if relevant:
                                Primary terms: {', '.join([term_info.get('term', '') for term_info in term_data.get('primary_terms', [])[:5]])}
                                
                                Format with proper HTML paragraph tags.
                            """}
                        ],
                        temperature=0.4
                    )
                    
                    subsection_content = subsection_content_response.content[0].text
                    optimized_sections.append(subsection_content)
                    
                    # Estimate added word count
                    subsection_words = len(re.findall(r'\b\w+\b', subsection_content))
                    current_word_count += subsection_words
                    
                    # Track change
                    structure_changes.append(f"Added subsection: {h3} under {h2}")
        
        # Add a conclusion if we have room
        if current_word_count < (target_word_count * 0.95):
            optimized_sections.append("<h2>Conclusion</h2>")
            
            conclusion_word_limit = min(200, target_word_count - current_word_count)
            
            conclusion_response = client.messages.create(
                model="claude-3-7-sonnet-20250219",
                max_tokens=conclusion_word_limit * 2,  # Allow some extra tokens for HTML
                system="You are an expert at writing concise, impactful conclusions.",
                messages=[
                    {"role": "user", "content": f"""
                        Write a brief conclusion for an article about "{keyword}".
                        
                        Requirements:
                        1. Summarize key points
                        2. Include a call to action
                        3. Reinforce the keyword relevance
                        4. Include at least 2 primary terms: {', '.join([term_info.get('term', '') for term_info in term_data.get('primary_terms', [])[:5]])}
                        5. STRICTLY limit to {conclusion_word_limit} words
                        
                        Format with proper HTML paragraph tags.
                    """}
                ],
                temperature=0.4
            )
            
            conclusion_content = conclusion_response.content[0].text
            optimized_sections.append(conclusion_content)
        
        # Create final document with change summary
        optimized_html = "\n".join(optimized_sections)
        
        # Create a more user-friendly change summary
        change_summary = f"""
        <div class="change-summary">
            <h2>Optimization Summary</h2>
            
            <p>This document has been optimized for the keyword <strong>"{keyword}"</strong> while preserving valuable original content.</p>
            
            <h3>Key Improvements:</h3>
            <ul>
                <li>Restructured content to better match search intent</li>
                <li>Enhanced keyword relevance throughout the document</li>
                <li>Created a focused article of approximately {current_word_count} words</li>
                <li>Added {len(content_additions)} new sections to address content gaps</li>
                <li>Incorporated important terms identified in top-ranking content</li>
            </ul>
            
            <h3>Structure Changes:</h3>
            <ul>
                {"".join(f"<li>{change}</li>" for change in structure_changes[:5])}
                {f"<li>Plus {len(structure_changes) - 5} additional structure changes</li>" if len(structure_changes) > 5 else ""}
            </ul>
            
            <h3>Content Enhancements:</h3>
            <ul>
                {"".join(f"<li>{improvement}</li>" for improvement in content_improvements[:5])}
                {f"<li>Plus {len(content_improvements) - 5} additional content improvements</li>" if len(content_improvements) > 5 else ""}
            </ul>
        </div>
        """
        
        return optimized_html, change_summary, True
        
    except Exception as e:
        error_msg = f"Exception in generate_optimized_article_with_tracking: {str(e)}"
        logger.error(error_msg)
        logger.error(traceback.format_exc())
        return "", "", False

def create_word_document_from_html(html_content: str, keyword: str, change_summary: str = "", 
                                  score_data: Dict = None) -> BytesIO:
    """
    Enhanced document creation with content score information
    Returns: document_stream
    """
    try:
        doc = Document()
        
        # Add document title
        title = doc.add_heading(f'Optimized Content: {keyword}', 0)
        title.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        
        # Add date
        date_para = doc.add_paragraph(f"Generated on: {time.strftime('%Y-%m-%d %H:%M:%S')}")
        date_para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        
        # Add horizontal line
        doc.add_paragraph("_" * 50)
        
        # Add content score if available
        if score_data:
            doc.add_heading("Content Score", 1)
            
            score_para = doc.add_paragraph()
            score_para.add_run(f"Overall Score: ").bold = True
            
            overall_score = score_data.get('overall_score', 0)
            grade = score_data.get('grade', 'F')
            
            score_run = score_para.add_run(f"{overall_score} ({grade})")
            
            # Color the score based on value
            if overall_score >= 70:
                score_run.font.color.rgb = RGBColor(0, 128, 0)  # Green
            elif overall_score < 50:
                score_run.font.color.rgb = RGBColor(255, 0, 0)  # Red
            else:
                score_run.font.color.rgb = RGBColor(255, 165, 0)  # Orange
            
            # Component scores
            components = score_data.get('components', {})
            for component, score in components.items():
                component_para = doc.add_paragraph(style='List Bullet')
                component_name = component.replace('_score', '').replace('_', ' ').title()
                component_para.add_run(f"{component_name}: ").bold = True
                component_para.add_run(f"{score}")
            
            # Add separator
            doc.add_paragraph("_" * 50)
        
        # Add change summary if provided
        if change_summary:
            doc.add_heading("Optimization Summary", 1)
            
            # Parse HTML summary
            summary_soup = BeautifulSoup(change_summary, 'html.parser')
            
            # Extract key points
            for h3 in summary_soup.find_all('h3'):
                doc.add_heading(h3.get_text(), 2)
                
                # Get the list that follows this heading
                ul = h3.find_next('ul')
                if ul:
                    for li in ul.find_all('li'):
                        doc.add_paragraph(li.get_text(), style='List Bullet')
            
            # Add separator before main content
            doc.add_paragraph()
            doc.add_paragraph("_" * 50)
            doc.add_paragraph()
        
        # Add content heading
        doc.add_heading("Optimized Content", 1)
        
        # IMPROVED: Direct, simpler approach to extract headings and content
        if html_content and isinstance(html_content, str):
            logger.info(f"Processing HTML content with length: {len(html_content)}")
            
            # Check if content contains HTML
            if '<' in html_content and '>' in html_content:
                # Parse HTML content with Beautiful Soup
                soup = BeautifulSoup(html_content, 'html.parser')
                
                # Log some debug information
                logger.info(f"HTML structure overview: {soup.prettify()[:500]}")
                
                # Direct extraction of headings and content, level by level
                # First, extract H1
                h1_headings = soup.find_all('h1')
                for h1 in h1_headings:
                    heading = doc.add_heading(h1.get_text().strip(), level=1)
                    heading.runs[0].font.size = Pt(16)
                    heading.runs[0].bold = True
                
                # Extract H2
                h2_headings = soup.find_all('h2')
                for h2 in h2_headings:
                    # Add the heading
                    heading = doc.add_heading(h2.get_text().strip(), level=2)
                    heading.runs[0].font.size = Pt(14)
                    heading.runs[0].bold = True
                    
                    # Find content between this H2 and the next heading
                    content_elements = []
                    next_element = h2.next_sibling
                    while next_element:
                        if hasattr(next_element, 'name') and next_element.name in ['h1', 'h2', 'h3', 'h4', 'h5', 'h6']:
                            break
                        if hasattr(next_element, 'name') and next_element.name in ['p', 'ul', 'ol']:
                            content_elements.append(next_element)
                        next_element = next_element.next_sibling
                    
                    # Add the content elements
                    for element in content_elements:
                        if element.name == 'p':
                            text = element.get_text().strip()
                            if text:
                                doc.add_paragraph(text)
                        elif element.name == 'ul':
                            for li in element.find_all('li'):
                                text = li.get_text().strip()
                                if text:
                                    doc.add_paragraph(text, style='List Bullet')
                        elif element.name == 'ol':
                            for li in element.find_all('li'):
                                text = li.get_text().strip()
                                if text:
                                    doc.add_paragraph(text, style='List Number')
                
                # Extract H3
                h3_headings = soup.find_all('h3')
                for h3 in h3_headings:
                    # Add the heading
                    heading = doc.add_heading(h3.get_text().strip(), level=3)
                    heading.runs[0].font.size = Pt(12)
                    heading.runs[0].bold = True
                    heading.runs[0].italic = True
                    
                    # Find content between this H3 and the next heading
                    content_elements = []
                    next_element = h3.next_sibling
                    while next_element:
                        if hasattr(next_element, 'name') and next_element.name in ['h1', 'h2', 'h3', 'h4', 'h5', 'h6']:
                            break
                        if hasattr(next_element, 'name') and next_element.name in ['p', 'ul', 'ol']:
                            content_elements.append(next_element)
                        next_element = next_element.next_sibling
                    
                    # Add the content elements
                    for element in content_elements:
                        if element.name == 'p':
                            text = element.get_text().strip()
                            if text:
                                doc.add_paragraph(text)
                        elif element.name == 'ul':
                            for li in element.find_all('li'):
                                text = li.get_text().strip()
                                if text:
                                    doc.add_paragraph(text, style='List Bullet')
                        elif element.name == 'ol':
                            for li in element.find_all('li'):
                                text = li.get_text().strip()
                                if text:
                                    doc.add_paragraph(text, style='List Number')
                
                # Extract H4
                h4_headings = soup.find_all('h4')
                for h4 in h4_headings:
                    # Add the heading
                    heading = doc.add_heading(h4.get_text().strip(), level=4)
                    heading.runs[0].font.size = Pt(11)
                    heading.runs[0].bold = True
                    heading.runs[0].italic = True
                    
                    # Find content between this H4 and the next heading
                    content_elements = []
                    next_element = h4.next_sibling
                    while next_element:
                        if hasattr(next_element, 'name') and next_element.name in ['h1', 'h2', 'h3', 'h4', 'h5', 'h6']:
                            break
                        if hasattr(next_element, 'name') and next_element.name in ['p', 'ul', 'ol']:
                            content_elements.append(next_element)
                        next_element = next_element.next_sibling
                    
                    # Add the content elements
                    for element in content_elements:
                        if element.name == 'p':
                            text = element.get_text().strip()
                            if text:
                                doc.add_paragraph(text)
                        elif element.name == 'ul':
                            for li in element.find_all('li'):
                                text = li.get_text().strip()
                                if text:
                                    doc.add_paragraph(text, style='List Bullet')
                        elif element.name == 'ol':
                            for li in element.find_all('li'):
                                text = li.get_text().strip()
                                if text:
                                    doc.add_paragraph(text, style='List Number')
                
                # If we didn't find any headings through the direct approach, try an alternative method
                if not (h1_headings or h2_headings or h3_headings or h4_headings):
                    logger.warning("No headings found with direct approach. Trying alternative method.")
                    
                    # Alternative: extract all text nodes and look for heading patterns
                    for tag in soup.find_all(['h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'p', 'ul', 'ol']):
                        if tag.name in ['h1', 'h2', 'h3', 'h4', 'h5', 'h6']:
                            level = int(tag.name[1])
                            text = tag.get_text().strip()
                            if text:
                                doc.add_heading(text, level=level)
                        elif tag.name == 'p':
                            text = tag.get_text().strip()
                            if text:
                                doc.add_paragraph(text)
                        elif tag.name == 'ul':
                            for li in tag.find_all('li'):
                                text = li.get_text().strip()
                                if text:
                                    doc.add_paragraph(text, style='List Bullet')
                        elif tag.name == 'ol':
                            for li in tag.find_all('li'):
                                text = li.get_text().strip()
                                if text:
                                    doc.add_paragraph(text, style='List Number')
            else:
                # Plain text handling
                logger.info("Content appears to be plain text. Processing line by line.")
                lines = html_content.split('\n')
                for line in lines:
                    line = line.strip()
                    if not line:
                        continue
                    
                    # Try to guess if this is a heading based on length and formatting
                    if len(line) < 80:
                        if line.startswith('#'):  # Markdown style
                            heading_level = len(re.match(r'^#+', line).group(0))
                            heading_text = line.lstrip('#').strip()
                            if heading_text:
                                doc.add_heading(heading_text, level=min(heading_level, 6))
                        elif line.isupper():  # ALL CAPS might be a heading
                            doc.add_heading(line, level=2)
                        elif re.match(r'^[A-Z][\w\s]+[:.?!]$', line):  # Title Case with ending punctuation
                            doc.add_heading(line, level=3)
                        else:  # Regular paragraph
                            doc.add_paragraph(line)
                    else:  # Regular paragraph
                        doc.add_paragraph(line)
        
        # Save document to memory stream
        doc_stream = BytesIO()
        doc.save(doc_stream)
        doc_stream.seek(0)
        
        return doc_stream
    
    except Exception as e:
        error_msg = f"Exception in create_word_document_from_html: {str(e)}"
        logger.error(error_msg)
        logger.error(traceback.format_exc())
        return BytesIO()

###############################################################################
# 12. Main Streamlit App
###############################################################################

def main():
    st.title("📊 SEO Content Optimizer")
    
    # Sidebar for API credentials
    st.sidebar.header("API Credentials")
    
    dataforseo_login = st.sidebar.text_input("DataForSEO API Login", type="password")
    dataforseo_password = st.sidebar.text_input("DataForSEO API Password", type="password")
    
    openai_api_key = st.sidebar.text_input("OpenAI API Key", type="password")
    anthropic_api_key = st.sidebar.text_input("Anthropic API Key", type="password")
    
    # Initialize session state
    if 'results' not in st.session_state:
        st.session_state.results = {}
    
    # Main content
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
        keyword = st.text_input("Target Keyword")
        
        if st.button("Fetch SERP Data"):
            if not keyword:
                st.error("Please enter a target keyword")
            elif not dataforseo_login or not dataforseo_password:
                st.error("Please enter DataForSEO API credentials")
            elif not openai_api_key or not anthropic_api_key:
                st.error("Please enter API keys for OpenAI and Anthropic")
            else:
                with st.spinner("Fetching SERP data..."):
                    # Fetch SERP results (updated function with PAA questions)
                    start_time = time.time()
                    organic_results, serp_features, paa_questions, serp_success = fetch_serp_results(
                        keyword, dataforseo_login, dataforseo_password
                    )
                    
                    if serp_success:
                        st.session_state.results['keyword'] = keyword
                        st.session_state.results['organic_results'] = organic_results
                        st.session_state.results['serp_features'] = serp_features
                        st.session_state.results['paa_questions'] = paa_questions
                        
                        # Show SERP results
                        st.subheader("Top 10 Organic Results")
                        df_results = pd.DataFrame(organic_results)
                        st.dataframe(df_results)
                        
                        # Show SERP features
                        st.subheader("SERP Features")
                        df_features = pd.DataFrame(serp_features)
                        st.dataframe(df_features)
                        
                        # Show People Also Asked questions
                        if paa_questions:
                            st.subheader("People Also Asked Questions")
                            for q in paa_questions:
                                st.write(f"- {q.get('question', '')}")
                        
                        # Fetch related keywords using DataForSEO
                        st.text("Fetching related keywords...")
                        related_keywords, kw_success = fetch_related_keywords_dataforseo(
                            keyword, dataforseo_login, dataforseo_password
                        )
                        
                        # Validate related keywords data
                        validated_keywords = []
                        for kw in related_keywords:
                            validated_kw = {
                                'keyword': kw.get('keyword', ''),
                                'search_volume': int(kw.get('search_volume', 0)) if kw.get('search_volume') is not None else 0,
                                'cpc': float(kw.get('cpc', 0.0)) if kw.get('cpc') is not None else 0.0,
                                'competition': float(kw.get('competition', 0.0)) if kw.get('competition') is not None else 0.0
                            }
                            validated_keywords.append(validated_kw)
                        
                        st.session_state.results['related_keywords'] = validated_keywords
                        
                        # Display related keywords
                        st.subheader("Related Keywords")
                        df_keywords = pd.DataFrame(validated_keywords)
                        st.dataframe(df_keywords)
                        
                        st.success(f"SERP analysis completed in {format_time(time.time() - start_time)}")
                    else:
                        st.error("Failed to fetch SERP data. Please check your API credentials.")
        
        # Show previously fetched data if available
        if 'organic_results' in st.session_state.results:
            st.subheader("Previously Fetched SERP Results")
            st.dataframe(pd.DataFrame(st.session_state.results['organic_results']))
            
            if 'related_keywords' in st.session_state.results:
                st.subheader("Previously Fetched Related Keywords")
                st.dataframe(pd.DataFrame(st.session_state.results['related_keywords']))
            
            if 'paa_questions' in st.session_state.results:
                st.subheader("Previously Fetched 'People Also Asked' Questions")
                for q in st.session_state.results['paa_questions']:
                    st.write(f"- {q.get('question', '')}")
    
    # Tab 2: Content Analysis
    with tabs[1]:
        st.header("Content Analysis")
        
        if 'organic_results' not in st.session_state.results:
            st.warning("Please fetch SERP data first (in the 'Input & SERP Analysis' tab)")
        else:
            if st.button("Analyze Content"):
                if not anthropic_api_key:
                    st.error("Please enter Anthropic API key")
                else:
                    with st.spinner("Analyzing content from top-ranking pages..."):
                        start_time = time.time()
                        
                        # Scrape and analyze content from top pages
                        scraped_contents = []
                        progress_bar = st.progress(0)
                        
                        for i, result in enumerate(st.session_state.results['organic_results']):
                            st.text(f"Scraping content from {result['url']}...")
                            content, success = scrape_webpage(result['url'])
                            
                            if success and content and content != "[Content not accessible due to site restrictions]":
                                scraped_contents.append({
                                    'url': result['url'],
                                    'title': result['title'],
                                    'content': content
                                })
                                
                                # Also extract headings
                                headings = extract_headings(result['url'])
                                scraped_contents[-1]['headings'] = headings
                            else:
                                st.warning(f"Could not scrape content from {result['url']}")
                            
                            progress_bar.progress((i + 1) / len(st.session_state.results['organic_results']))
                        
                        if not scraped_contents:
                            st.error("Could not scrape content from any URLs. Please try a different keyword.")
                            return
                            
                        st.session_state.results['scraped_contents'] = scraped_contents
                        
                        # Analyze semantic structure using Claude
                        st.text("Analyzing semantic structure...")
                        semantic_structure, structure_success = analyze_semantic_structure(
                            scraped_contents, anthropic_api_key
                        )
                        
                        if structure_success:
                            st.session_state.results['semantic_structure'] = semantic_structure
                            
                            # Extract important terms using Claude
                            st.text("Extracting important terms and topics...")
                            term_data, term_success = extract_important_terms(
                                scraped_contents, anthropic_api_key
                            )
                            
                            if term_success:
                                st.session_state.results['term_data'] = term_data
                            
                            st.subheader("Recommended Semantic Structure")
                            st.write(f"**H1:** {semantic_structure.get('h1', '')}")
                            
                            for i, section in enumerate(semantic_structure.get('sections', []), 1):
                                st.write(f"**H2 {i}:** {section.get('h2', '')}")
                                
                                for j, subsection in enumerate(section.get('subsections', []), 1):
                                    st.write(f"  - **H3 {i}.{j}:** {subsection.get('h3', '')}")
                            
                            if term_success:
                                st.subheader("Important Terms")
                                
                                # Display primary terms
                                if 'primary_terms' in term_data:
                                    st.write("**Primary Terms:**")
                                    primary_df = pd.DataFrame(term_data['primary_terms'])
                                    st.dataframe(primary_df)
                                
                                # Display secondary terms
                                if 'secondary_terms' in term_data:
                                    st.write("**Secondary Terms:**")
                                    secondary_df = pd.DataFrame(term_data['secondary_terms'])
                                    st.dataframe(secondary_df)
                            
                            st.success(f"Content analysis completed in {format_time(time.time() - start_time)}")
                        else:
                            st.error("Failed to analyze semantic structure")
            
            # Show previously analyzed structure if available
            if 'semantic_structure' in st.session_state.results:
                st.subheader("Previously Analyzed Semantic Structure")
                semantic_structure = st.session_state.results['semantic_structure']
                
                st.write(f"**H1:** {semantic_structure.get('h1', '')}")
                
                for i, section in enumerate(semantic_structure.get('sections', []), 1):
                    st.write(f"**H2 {i}:** {section.get('h2', '')}")
                    
                    for j, subsection in enumerate(section.get('subsections', []), 1):
                        st.write(f"  - **H3 {i}.{j}:** {subsection.get('h3', '')}")
                
                # Show previously extracted term data if available
                if 'term_data' in st.session_state.results:
                    with st.expander("View Extracted Terms & Topics"):
                        term_data = st.session_state.results['term_data']
                        
                        # Display primary terms
                        if 'primary_terms' in term_data:
                            st.write("**Primary Terms:**")
                            primary_df = pd.DataFrame(term_data['primary_terms'])
                            st.dataframe(primary_df)
                        
                        # Display secondary terms
                        if 'secondary_terms' in term_data:
                            st.write("**Secondary Terms:**")
                            secondary_df = pd.DataFrame(term_data['secondary_terms'])
                            st.dataframe(secondary_df)
                        
                        # Display topics
                        if 'topics' in term_data:
                            st.write("**Topics to Cover:**")
                            topics_df = pd.DataFrame(term_data['topics'])
                            st.dataframe(topics_df)
    
    # Tab 3: Article Generation
    with tabs[2]:
        st.header("Article Generation")
        
        if 'semantic_structure' not in st.session_state.results:
            st.warning("Please complete content analysis first (in the 'Content Analysis' tab)")
        else:
            # Add option for guidance-only mode
            content_type = st.radio(
                "Content Generation Type:",
                ["Full Article", "Writing Guidance Only"],
                help="Choose 'Full Article' for complete content or 'Writing Guidance Only' for section-by-section writing directions"
            )
            guidance_only = (content_type == "Writing Guidance Only")
            
            if st.button("Generate " + ("Content Guidance" if guidance_only else "Article") + " and Meta Tags"):
                if not anthropic_api_key:
                    st.error("Please enter Anthropic API key")
                else:
                    with st.spinner("Generating " + ("content guidance" if guidance_only else "article") + " and meta tags..."):
                        start_time = time.time()
                        
                        # Use term data if available
                        term_data = st.session_state.results.get('term_data', {})
                        
                        # Generate article or guidance using Claude
                        article_content, article_success = generate_article(
                            st.session_state.results['keyword'],
                            st.session_state.results['semantic_structure'],
                            st.session_state.results.get('related_keywords', []),
                            st.session_state.results.get('serp_features', []),
                            st.session_state.results.get('paa_questions', []),
                            term_data,
                            anthropic_api_key,
                            guidance_only
                        )
                        
                        if article_success and article_content:
                            # Store with special key for guidance
                            if guidance_only:
                                st.session_state.results['guidance_content'] = article_content
                            else:
                                st.session_state.results['article_content'] = article_content
                            
                            # Store the guidance flag
                            st.session_state.results['guidance_only'] = guidance_only
                            
                            # Generate meta title and description with Claude
                            meta_title, meta_description, meta_success = generate_meta_tags(
                                st.session_state.results['keyword'],
                                st.session_state.results['semantic_structure'],
                                st.session_state.results.get('related_keywords', []),
                                term_data,
                                anthropic_api_key
                            )
                            
                            if meta_success:
                                st.session_state.results['meta_title'] = meta_title
                                st.session_state.results['meta_description'] = meta_description
                                
                                st.subheader("Meta Tags")
                                st.write(f"**Meta Title:** {meta_title}")
                                st.write(f"**Meta Description:** {meta_description}")
                            
                            # Score the content if it's a full article
                            if not guidance_only and 'term_data' in st.session_state.results:
                                try:
                                    score_data, score_success = score_content(
                                        article_content,
                                        st.session_state.results['term_data'],
                                        st.session_state.results['keyword']
                                    )
                                    
                                    if score_success:
                                        st.session_state.results['content_score'] = score_data
                                        
                                        # Display content score
                                        st.subheader("Content Score")
                                        score = score_data.get('overall_score', 0)
                                        grade = score_data.get('grade', 'F')
                                        
                                        # CSS to style the score display
                                        score_color = "green" if score >= 70 else "red" if score < 50 else "orange"
                                        st.markdown(f"""
                                        <div style="background-color: #f0f0f0; padding: 10px; border-radius: 5px;">
                                            <h3 style="margin:0;">Content Score: <span style="color:{score_color};">{score} ({grade})</span></h3>
                                        </div>
                                        """, unsafe_allow_html=True)
                                except Exception as e:
                                    logger.error(f"Error scoring content: {e}")
                            
                            st.subheader("Generated " + ("Content Guidance" if guidance_only else "Article"))
                            st.markdown(article_content, unsafe_allow_html=True)
                            
                            st.success(f"Content generation completed in {format_time(time.time() - start_time)}")
                        else:
                            st.error("Failed to generate " + ("content guidance" if guidance_only else "article") + ". Please try again.")
            
            # Show previously generated article or guidance if available
            if 'article_content' in st.session_state.results or 'guidance_content' in st.session_state.results:
                if 'meta_title' in st.session_state.results:
                    st.subheader("Previously Generated Meta Tags")
                    st.write(f"**Meta Title:** {st.session_state.results['meta_title']}")
                    st.write(f"**Meta Description:** {st.session_state.results['meta_description']}")
                
                # Display content score if available
                if 'content_score' in st.session_state.results:
                    st.subheader("Content Score")
                    score_data = st.session_state.results['content_score']
                    score = score_data.get('overall_score', 0)
                    grade = score_data.get('grade', 'F')
                    
                    # CSS to style the score display
                    score_color = "green" if score >= 70 else "red" if score < 50 else "orange"
                    st.markdown(f"""
                    <div style="background-color: #f0f0f0; padding: 10px; border-radius: 5px;">
                        <h3 style="margin:0;">Content Score: <span style="color:{score_color};">{score} ({grade})</span></h3>
                    </div>
                    """, unsafe_allow_html=True)
                
                # Display appropriate content based on what's available
                if 'guidance_only' in st.session_state.results and st.session_state.results['guidance_only']:
                    st.subheader("Previously Generated Content Guidance")
                    if 'guidance_content' in st.session_state.results:
                        st.markdown(st.session_state.results['guidance_content'], unsafe_allow_html=True)
                else:
                    st.subheader("Previously Generated Article")
                    if 'article_content' in st.session_state.results:
                        st.markdown(st.session_state.results['article_content'], unsafe_allow_html=True)
    
    # Tab 4: Internal Linking
    with tabs[3]:
        st.header("Internal Linking")
        
        is_guidance_only = st.session_state.results.get('guidance_only', False)
        
        # Check if content exists and determine which type
        has_content = False
        if is_guidance_only and 'guidance_content' in st.session_state.results:
            has_content = True
            content_type = "guidance"
        elif not is_guidance_only and 'article_content' in st.session_state.results:
            has_content = True
            content_type = "article"
        
        if not has_content:
            st.warning("Please generate content first (in the 'Article Generation' tab)")
        elif is_guidance_only:
            st.warning("Internal linking is only available for full articles, not writing guidance")
        else:
            st.write("Upload a spreadsheet with your site pages (CSV or Excel):")
            st.write("The spreadsheet must contain columns: URL, Title, Meta Description")
            
            # Create sample template button
            if st.button("Generate Sample Template"):
                # Create sample dataframe
                sample_data = {
                    'URL': ['https://example.com/page1', 'https://example.com/page2'],
                    'Title': ['Example Page 1', 'Example Page 2'],
                    'Meta Description': ['Description for page 1', 'Description for page 2']
                }
                sample_df = pd.DataFrame(sample_data)
                
                # Convert to CSV
                csv = sample_df.to_csv(index=False)
                
                # Provide download button
                st.download_button(
                    label="Download Sample CSV Template",
                    data=csv,
                    file_name="site_pages_template.csv",
                    mime="text/csv",
                )
            
            # File uploader for spreadsheet
            pages_file = st.file_uploader("Upload Site Pages Spreadsheet", type=['csv', 'xlsx', 'xls'])
            
            # Batch size for embedding
            batch_size = st.slider("Embedding Batch Size", 5, 50, 20, 
                                   help="Larger batch size is faster but may hit API limits")
            
            if st.button("Generate Internal Links"):
                if not openai_api_key:
                    st.error("Please enter OpenAI API key for embeddings")
                elif not anthropic_api_key:
                    st.error("Please enter Anthropic API key for anchor text selection")
                elif not pages_file:
                    st.error("Please upload a spreadsheet with site pages")
                else:
                    with st.spinner("Processing site pages and generating internal links..."):
                        start_time = time.time()
                        
                        # Parse spreadsheet
                        pages, parse_success = parse_site_pages_spreadsheet(pages_file)
                        
                        if parse_success and pages:
                            # Status update
                            status_text = st.empty()
                            status_text.text(f"Generating embeddings for {len(pages)} site pages...")
                            
                            # Generate embeddings for site pages using OpenAI
                            pages_with_embeddings, embed_success = embed_site_pages(
                                pages, openai_api_key, batch_size
                            )
                            
                            if embed_success:
                                # Count words in the article
                                article_content = st.session_state.results['article_content']
                                word_count = len(re.findall(r'\w+', article_content))
                                
                                status_text.text(f"Analyzing article content and generating internal links...")
                                
                                # Generate internal links using OpenAI for embeddings and Claude for anchor text
                                article_with_links, links_added, links_success = generate_internal_links_with_embeddings(
                                    article_content, pages_with_embeddings, anthropic_api_key, word_count
                                )
                                
                                if links_success:
                                    st.session_state.results['article_with_links'] = article_with_links
                                    st.session_state.results['internal_links'] = links_added
                                    
                                    st.subheader("Article With Internal Links")
                                    st.markdown(article_with_links, unsafe_allow_html=True)
                                    
                                    st.subheader("Internal Links Added")
                                    for link in links_added:
                                        st.write(f"**URL:** {link.get('url')}")
                                        st.write(f"**Anchor Text:** {link.get('anchor_text')}")
                                        st.write(f"**Context:** {link.get('context')}")
                                        st.write("---")
                                    
                                    st.success(f"Internal linking completed in {format_time(time.time() - start_time)}")
                                else:
                                    st.error("Failed to generate internal links")
                            else:
                                st.error("Failed to generate embeddings for site pages")
                        else:
                            st.error("Failed to parse spreadsheet or no pages found")
            
            # Show previously generated links if available
            if 'article_with_links' in st.session_state.results:
                st.subheader("Previously Generated Article With Internal Links")
                st.markdown(st.session_state.results['article_with_links'], unsafe_allow_html=True)
    
    # Tab 5: SEO Brief
    with tabs[4]:
        st.header("SEO Brief & Downloadable Report")
        
        # Check if content exists and determine which type
        has_content = False
        content_type = "none"
        
        if 'guidance_only' in st.session_state.results:
            is_guidance = st.session_state.results['guidance_only']
            if is_guidance and 'guidance_content' in st.session_state.results:
                has_content = True
                content_type = "guidance"
            elif not is_guidance and 'article_content' in st.session_state.results:
                has_content = True
                content_type = "article"
        
        if not has_content:
            st.warning("Please generate content first (in the 'Article Generation' tab)")
        else:
            if st.button("Generate SEO Brief"):
                with st.spinner("Generating SEO brief..."):
                    start_time = time.time()
                    
                    # Determine which content to use
                    if content_type == "article":
                        # Use article with internal links if available, otherwise use regular article
                        article_content = st.session_state.results.get('article_with_links', 
                                                                      st.session_state.results['article_content'])
                        
                        internal_links = st.session_state.results.get('internal_links', None)
                        guidance_only = False
                    else:
                        # Use guidance content
                        article_content = st.session_state.results['guidance_content']
                        internal_links = None
                        guidance_only = True
                    
                    # Get meta title and description
                    meta_title = st.session_state.results.get('meta_title', 
                                                             f"{st.session_state.results['keyword']} - Complete Guide")
                    
                    meta_description = st.session_state.results.get('meta_description', 
                                                                  f"Learn everything about {st.session_state.results['keyword']} in our comprehensive guide.")
                    
                    # Get PAA questions
                    paa_questions = st.session_state.results.get('paa_questions', [])
                    
                    # Get term data and content score if available
                    term_data = st.session_state.results.get('term_data', None)
                    score_data = st.session_state.results.get('content_score', None)
                    
                    # Create Word document with enhanced Claude content
                    doc_stream, doc_success = create_word_document(
                        st.session_state.results['keyword'],
                        st.session_state.results['organic_results'],
                        st.session_state.results.get('related_keywords', []),
                        st.session_state.results['semantic_structure'],
                        article_content,
                        meta_title,
                        meta_description,
                        paa_questions,
                        term_data,
                        score_data,
                        internal_links,
                        guidance_only
                    )
                    
                    if doc_success:
                        st.session_state.results['doc_stream'] = doc_stream
                        
                        # First download button
                        st.download_button(
                            label="Download SEO Brief Now",
                            data=doc_stream,
                            file_name=f"seo_brief_{st.session_state.results['keyword'].replace(' ', '_')}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            key="download_brief_1"  # Added unique key
                        )
                        
                        st.success(f"SEO brief generation completed in {format_time(time.time() - start_time)}")
                    else:
                        st.error("Failed to generate SEO brief document")
            
            # Show summary of all components
            st.subheader("SEO Analysis Summary")
            
            components = [
                ("Keyword", 'keyword' in st.session_state.results),
                ("SERP Analysis", 'organic_results' in st.session_state.results),
                ("People Also Asked", 'paa_questions' in st.session_state.results),
                ("Related Keywords", 'related_keywords' in st.session_state.results),
                ("Content Analysis", 'scraped_contents' in st.session_state.results),
                ("Term Analysis", 'term_data' in st.session_state.results),
                ("Semantic Structure", 'semantic_structure' in st.session_state.results),
                ("Meta Title & Description", 'meta_title' in st.session_state.results),
                ("Content Score", 'content_score' in st.session_state.results),
                ("Generated Content", 'article_content' in st.session_state.results or 'guidance_content' in st.session_state.results),
                ("Internal Linking", 'article_with_links' in st.session_state.results)
            ]
            
            for component, status in components:
                st.write(f"**{component}:** {'✅ Completed' if status else '❌ Not Completed'}")
            
            # Display download button if available with different key
            if 'doc_stream' in st.session_state.results:
                st.subheader("Download SEO Brief")
                st.download_button(
                    label="Download SEO Brief Document",  # Changed label to avoid duplicate element ID
                    data=st.session_state.results['doc_stream'],
                    file_name=f"seo_brief_{st.session_state.results['keyword'].replace(' ', '_')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key="download_brief_2"  # Added unique key
                )
    
    # Tab 6: Content Updates
    with tabs[5]:
        st.header("Content Update Recommendations")
        
        if 'semantic_structure' not in st.session_state.results or 'scraped_contents' not in st.session_state.results:
            st.warning("Please complete SERP and content analysis first (in the 'Input & SERP Analysis' and 'Content Analysis' tabs)")
        else:
            st.write("Upload your existing content document to get update recommendations based on competitor analysis:")
            
            # File uploader for document
            content_file = st.file_uploader("Upload Content Document", type=['docx'])
            
            # Add radio button BEFORE the button click event
            update_type = st.radio(
                "Select update approach:",
                ["Recommendations Only", "Generate Optimized Article"],
                help="Choose whether to receive recommendations or get a completely optimized article"
            )
            
            if st.button("Generate Content Updates"):
                if not anthropic_api_key:
                    st.error("Please enter Anthropic API key")
                elif not content_file:
                    st.error("Please upload a content document")
                else:
                    with st.spinner("Analyzing content and generating updates..."):
                        start_time = time.time()
                        
                        # Parse uploaded document
                        existing_content, parse_success = parse_word_document(content_file)
                        
                        if parse_success and existing_content:
                            # Get term data and score data if available
                            term_data = st.session_state.results.get('term_data', {})
                            
                            # Score the existing content if we have term data
                            score_data = None
                            if term_data:
                                html_content = f"<p>{existing_content.get('full_text', '').replace('\n\n', '</p><p>')}</p>"
                                score_data, score_success = score_content(
                                    html_content,
                                    term_data,
                                    st.session_state.results['keyword']
                                )
                                
                                if score_success:
                                    st.session_state.results['existing_content_score'] = score_data
                            
                            # Enhanced content gap analysis with Claude
                            content_gaps, gap_success = analyze_content_gaps(
                                existing_content,
                                st.session_state.results['scraped_contents'],
                                st.session_state.results['semantic_structure'],
                                term_data,
                                score_data if score_data else {},
                                anthropic_api_key,
                                st.session_state.results['keyword'],
                                st.session_state.results.get('paa_questions', [])
                            )
                            
                            if gap_success and content_gaps:
                                # Store results
                                st.session_state.results['existing_content'] = existing_content
                                st.session_state.results['content_gaps'] = content_gaps
                                
                                if update_type == "Recommendations Only":
                                    # Create updated document with recommendations
                                    updated_doc, doc_success = create_updated_document(
                                        existing_content,
                                        content_gaps,
                                        st.session_state.results['keyword'],
                                        score_data
                                    )
                                    
                                    if doc_success:
                                        st.session_state.results['updated_doc'] = updated_doc
                                        
                                        # Display content score if available
                                        if score_data:
                                            st.subheader("Content Score Assessment")
                                            score = score_data.get('overall_score', 0)
                                            grade = score_data.get('grade', 'F')
                                            
                                            # CSS to style the score display
                                            score_color = "green" if score >= 70 else "red" if score < 50 else "orange"
                                            st.markdown(f"""
                                            <div style="background-color: #f0f0f0; padding: 10px; border-radius: 5px;">
                                                <h3 style="margin:0;">Content Score: <span style="color:{score_color};">{score} ({grade})</span></h3>
                                                <p>Implementing recommendations should significantly improve this score.</p>
                                            </div>
                                            """, unsafe_allow_html=True)
                                        
                                        # Display summary of recommendations
                                        st.subheader("Content Update Recommendations")
                                        
                                        # Term usage issues (new section)
                                        if content_gaps.get('term_usage_issues'):
                                            st.markdown("### Term Usage Issues")
                                            term_table = pd.DataFrame(content_gaps['term_usage_issues'])
                                            st.dataframe(term_table)
                                        
                                        # Semantic relevancy issues
                                        if content_gaps.get('semantic_relevancy_issues'):
                                            st.markdown("### Semantic Relevancy Issues")
                                            for issue in content_gaps['semantic_relevancy_issues']:
                                                st.markdown(f"**{issue.get('section', '')}**")
                                                st.markdown(f"*Issue:* {issue.get('issue', '')}")
                                                st.markdown(f"*Recommendation:* **<span style='color:red'>{issue.get('recommendation', '')}</span>**", unsafe_allow_html=True)
                                        
                                        # Unanswered PAA questions
                                        if content_gaps.get('unanswered_questions'):
                                            st.markdown("### Unanswered 'People Also Asked' Questions")
                                            for question in content_gaps['unanswered_questions']:
                                                st.markdown(f"**Q: {question.get('question', '')}**")
                                                st.markdown(f"*Section:* {question.get('insert_into_section', 'END')}")
                                                st.markdown(f"*Suggested Answer:* **<span style='color:red'>{question.get('suggested_answer', '')}</span>**", unsafe_allow_html=True)
                                        
                                        # Missing headings
                                        if content_gaps.get('missing_headings'):
                                            st.markdown("### Missing Sections")
                                            for heading in content_gaps['missing_headings']:
                                                st.markdown(f"**{heading.get('heading', '')}**")
                                                st.markdown(f"*Insert after:* {heading.get('insert_after', 'END')}")
                                                st.markdown(f"*{heading.get('suggested_content', '')}*")
                                        
                                        # Revised headings
                                        if content_gaps.get('revised_headings'):
                                            st.markdown("### Heading Revisions")
                                            for revision in content_gaps['revised_headings']:
                                                st.markdown(f"~~{revision.get('original', '')}~~ → **<span style='color:red'>{revision.get('suggested', '')}</span>**", unsafe_allow_html=True)
                                        
                                        # Content gaps
                                        if content_gaps.get('content_gaps'):
                                            st.markdown("### Content Gaps")
                                            for gap in content_gaps['content_gaps']:
                                                st.markdown(f"**{gap.get('topic', '')}**: {gap.get('details', '')}")
                                        
                                        # Expansion areas
                                        if content_gaps.get('expansion_areas'):
                                            st.markdown("### Expansion Areas")
                                            for area in content_gaps['expansion_areas']:
                                                st.markdown(f"**{area.get('section', '')}**: {area.get('reason', '')}")
                                        
                                        # Download button
                                        st.download_button(
                                            label="Download Update Recommendations",
                                            data=updated_doc,
                                            file_name=f"content_updates_{st.session_state.results['keyword'].replace(' ', '_')}.docx",
                                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                                        )
                                        
                                        st.success(f"Content update recommendations generated in {format_time(time.time() - start_time)}")
                                    else:
                                        st.error("Failed to create updated document")
                                
                                else:  # Generate Optimized Article
                                    # Generate optimized article with Claude
                                    optimized_content, change_summary, success = generate_optimized_article_with_tracking(
                                        existing_content,
                                        st.session_state.results['scraped_contents'],
                                        st.session_state.results['semantic_structure'],
                                        st.session_state.results.get('related_keywords', []),
                                        st.session_state.results['keyword'],
                                        st.session_state.results.get('paa_questions', []),
                                        term_data,
                                        anthropic_api_key
                                    )
                                    
                                    if success and optimized_content:
                                        # Store the optimized content
                                        st.session_state.results['optimized_content'] = optimized_content
                                        st.session_state.results['change_summary'] = change_summary
                                        
                                        # Score the optimized content
                                        if term_data:
                                            optimized_score_data, score_success = score_content(
                                                optimized_content,
                                                term_data,
                                                st.session_state.results['keyword']
                                            )
                                            
                                            if score_success:
                                                st.session_state.results['optimized_content_score'] = optimized_score_data
                                                
                                                # Compare scores if we have both
                                                if score_data:
                                                    old_score = score_data.get('overall_score', 0)
                                                    new_score = optimized_score_data.get('overall_score', 0)
                                                    
                                                    st.subheader("Content Score Improvement")
                                                    
                                                    # Score comparison with styling
                                                    score_diff = new_score - old_score
                                                    st.markdown(f"""
                                                    <div style="background-color: #f0f0f0; padding: 15px; border-radius: 5px; margin-bottom: 20px;">
                                                        <div style="display: flex; justify-content: space-between; margin-bottom: 10px;">
                                                            <div>
                                                                <span style="font-size: 16px; font-weight: bold;">Original Score:</span>
                                                                <span style="font-size: 18px; color: {'green' if old_score >= 70 else 'red' if old_score < 50 else 'orange'}; font-weight: bold; margin-left: 10px;">
                                                                    {old_score} ({score_data.get('grade', 'F')})
                                                                </span>
                                                            </div>
                                                            <div>
                                                                <span style="font-size: 16px; font-weight: bold;">New Score:</span>
                                                                <span style="font-size: 18px; color: {'green' if new_score >= 70 else 'red' if new_score < 50 else 'orange'}; font-weight: bold; margin-left: 10px;">
                                                                    {new_score} ({optimized_score_data.get('grade', 'F')})
                                                                </span>
                                                            </div>
                                                        </div>
                                                        <div style="text-align: center; padding-top: 5px;">
                                                            <span style="font-size: 16px;">Improvement:</span>
                                                            <span style="font-size: 20px; color: {'green' if score_diff > 0 else 'red' if score_diff < 0 else 'gray'}; font-weight: bold; margin-left: 10px;">
                                                                {'+' if score_diff > 0 else ''}{score_diff} points
                                                            </span>
                                                        </div>
                                                    </div>
                                                    """, unsafe_allow_html=True)
                                        
                                        # Display a simplified tabbed view
                                        opt_tabs = st.tabs(["Optimized Article", "Optimization Summary"])
                                        
                                        with opt_tabs[0]:
                                            st.markdown("## Optimized Article")
                                            st.markdown(optimized_content, unsafe_allow_html=True)
                                        
                                        with opt_tabs[1]:
                                            st.markdown("## Optimization Summary")
                                            st.markdown(change_summary, unsafe_allow_html=True)
                                        
                                        # Create Word document from HTML with score data
                                        doc_stream = create_word_document_from_html(
                                            optimized_content, 
                                            st.session_state.results['keyword'],
                                            change_summary,
                                            st.session_state.results.get('optimized_content_score')
                                        )
                                        
                                        # Download button
                                        st.download_button(
                                            label="Download Optimized Article",
                                            data=doc_stream,
                                            file_name=f"optimized_{st.session_state.results['keyword'].replace(' ', '_')}.docx",
                                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                                        )
                                        
                                        st.success(f"Optimized article generated in {format_time(time.time() - start_time)}")
                                    else:
                                        st.error("Failed to generate optimized article")
                                        
            # Show previously generated recommendations if available
            if 'content_gaps' in st.session_state.results:
                if 'optimized_content' in st.session_state.results:
                    st.subheader("Previously Generated Optimized Article")
                    
                    # Display score improvement if available
                    if 'existing_content_score' in st.session_state.results and 'optimized_content_score' in st.session_state.results:
                        old_score = st.session_state.results['existing_content_score'].get('overall_score', 0)
                        new_score = st.session_state.results['optimized_content_score'].get('overall_score', 0)
                        
                        score_diff = new_score - old_score
                        st.markdown(f"""
                        <div style="background-color: #f0f0f0; padding: 15px; border-radius: 5px; margin-bottom: 20px;">
                            <div style="display: flex; justify-content: space-between; margin-bottom: 10px;">
                                <div>
                                    <span style="font-size: 16px; font-weight: bold;">Original Score:</span>
                                    <span style="font-size: 18px; color: {'green' if old_score >= 70 else 'red' if old_score < 50 else 'orange'}; font-weight: bold; margin-left: 10px;">
                                        {old_score} ({st.session_state.results['existing_content_score'].get('grade', 'F')})
                                    </span>
                                </div>
                                <div>
                                    <span style="font-size: 16px; font-weight: bold;">New Score:</span>
                                    <span style="font-size: 18px; color: {'green' if new_score >= 70 else 'red' if new_score < 50 else 'orange'}; font-weight: bold; margin-left: 10px;">
                                        {new_score} ({st.session_state.results['optimized_content_score'].get('grade', 'F')})
                                    </span>
                                </div>
                            </div>
                            <div style="text-align: center; padding-top: 5px;">
                                <span style="font-size: 16px;">Improvement:</span>
                                <span style="font-size: 20px; color: {'green' if score_diff > 0 else 'red' if score_diff < 0 else 'gray'}; font-weight: bold; margin-left: 10px;">
                                    {'+' if score_diff > 0 else ''}{score_diff} points
                                </span>
                            </div>
                        </div>
                        """, unsafe_allow_html=True)
                    
                    # If we have a previously generated change summary, show tabs
                    if 'change_summary' in st.session_state.results:
                        prev_tabs = st.tabs(["Optimized Article", "Change Summary"])
                        
                        with prev_tabs[0]:
                            st.markdown(st.session_state.results['optimized_content'], unsafe_allow_html=True)
                            
                        with prev_tabs[1]:
                            st.markdown(st.session_state.results['change_summary'], unsafe_allow_html=True)
                    else:
                        # Just show the content without tabs if no change summary
                        st.markdown(st.session_state.results['optimized_content'], unsafe_allow_html=True)
                    
                    # If we have a previously optimized article, offer download
                    if 'keyword' in st.session_state.results:
                        doc_stream = create_word_document_from_html(
                            st.session_state.results['optimized_content'],
                            st.session_state.results['keyword'],
                            st.session_state.results.get('change_summary', ''),
                            st.session_state.results.get('optimized_content_score')
                        )
                        
                        st.download_button(
                            label="Download Previous Optimized Article",
                            data=doc_stream,
                            file_name=f"optimized_{st.session_state.results['keyword'].replace(' ', '_')}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            key="download_previous_optimized"
                        )
                
                elif 'updated_doc' in st.session_state.results:
                    st.subheader("Previously Generated Update Recommendations")
                    
                    # Display content score if available
                    if 'existing_content_score' in st.session_state.results:
                        score_data = st.session_state.results['existing_content_score']
                        score = score_data.get('overall_score', 0)
                        grade = score_data.get('grade', 'F')
                        
                        score_color = "green" if score >= 70 else "red" if score < 50 else "orange"
                        st.markdown(f"""
                        <div style="background-color: #f0f0f0; padding: 10px; border-radius: 5px;">
                            <h3 style="margin:0;">Content Score: <span style="color:{score_color};">{score} ({grade})</span></h3>
                            <p>Implementing recommendations should significantly improve this score.</p>
                        </div>
                        """, unsafe_allow_html=True)
                    
                    # Download button for previously generated document
                    st.download_button(
                        label="Download Previous Update Recommendations",
                        data=st.session_state.results['updated_doc'],
                        file_name=f"content_updates_{st.session_state.results['keyword'].replace(' ', '_')}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key="download_previous_updates"
                    )
    
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
