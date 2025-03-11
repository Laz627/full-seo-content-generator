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

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# Set page configuration
st.set_page_config(
    page_title="SEO Analysis & Content Generator",
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

def create_default_keywords(keyword: str) -> List[Dict]:
    """Create default related keywords when API fails"""
    return [
        {'keyword': f"{keyword} guide", 'search_volume': 100, 'cpc': 0.5, 'competition': 0.3},
        {'keyword': f"best {keyword}", 'search_volume': 90, 'cpc': 0.6, 'competition': 0.4},
        {'keyword': f"{keyword} tutorial", 'search_volume': 80, 'cpc': 0.4, 'competition': 0.3},
        {'keyword': f"how to use {keyword}", 'search_volume': 70, 'cpc': 0.3, 'competition': 0.2},
        {'keyword': f"{keyword} benefits", 'search_volume': 60, 'cpc': 0.5, 'competition': 0.3},
        {'keyword': f"{keyword} examples", 'search_volume': 50, 'cpc': 0.4, 'competition': 0.2},
        {'keyword': f"{keyword} tips", 'search_volume': 40, 'cpc': 0.3, 'competition': 0.2},
        {'keyword': f"free {keyword}", 'search_volume': 100, 'cpc': 0.7, 'competition': 0.5},
        {'keyword': f"{keyword} alternatives", 'search_volume': 80, 'cpc': 0.8, 'competition': 0.6},
        {'keyword': f"{keyword} vs", 'search_volume': 90, 'cpc': 0.9, 'competition': 0.7}
    ]

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
                return create_default_keywords(keyword), False
            
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
                return create_default_keywords(keyword), False
        else:
            error_msg = f"HTTP Error: {response.status_code} - {response.text}"
            logger.error(error_msg)
            return create_default_keywords(keyword), False
    
    except Exception as e:
        error_msg = f"Exception in fetch_keyword_suggestions: {str(e)}"
        logger.error(error_msg)
        logger.error(traceback.format_exc())
        return create_default_keywords(keyword), False

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
# 5. Meta Title and Description Generation
###############################################################################

def generate_meta_tags(keyword: str, semantic_structure: Dict, related_keywords: List[Dict], 
                      openai_api_key: str) -> Tuple[str, str, bool]:
    """
    Generate meta title and description for the content
    Returns: meta_title, meta_description, success_status
    """
    try:
        openai.api_key = openai_api_key
        
        # Extract H1 and first few sections for context
        h1 = semantic_structure.get('h1', f"Complete Guide to {keyword}")
        
        # Get top 5 related keywords
        top_keywords = ", ".join([kw.get('keyword', '') for kw in related_keywords[:5] if kw.get('keyword')])
        
        # Generate meta tags
        response = openai.ChatCompletion.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": "You are an SEO specialist who creates optimized meta tags."},
                {"role": "user", "content": f"""
                Create an SEO-optimized meta title and description for an article about "{keyword}".
                
                The article's main heading is: "{h1}"
                
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
        content = response.choices[0].message.content
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
# 6. Embeddings and Semantic Analysis
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

def analyze_semantic_structure(contents: List[Dict], openai_api_key: str) -> Tuple[Dict, bool]:
    """
    Analyze semantic structure of content to determine optimal hierarchy
    Returns: semantic_analysis, success_status
    """
    try:
        openai.api_key = openai_api_key
        
        # Combine all content for analysis
        combined_content = "\n\n".join([c.get('content', '') for c in contents if c.get('content')])
        
        # Prepare summarized content if it's too long
        if len(combined_content) > 10000:
            combined_content = combined_content[:10000]
        
        # Use OpenAI to analyze content and suggest headings structure
        response = openai.ChatCompletion.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": "You are an SEO expert specializing in content structure."},
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
        content = response.choices[0].message.content
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
# 7. Content Generation
###############################################################################

def verify_semantic_match(anchor_text: str, page_title: str) -> bool:
    """
    Verify that anchor text and page title have meaningful word overlap after removing stop words
    """
    # Define common stop words
    stop_words = {'a', 'an', 'the', 'and', 'or', 'but', 'in', 'on', 'at', 'to', 'for', 'with', 
                  'by', 'about', 'as', 'is', 'are', 'was', 'were', 'of', 'from', 'into', 'during',
                  'after', 'before', 'above', 'below', 'between', 'under', 'over', 'through', 'it',
                  'its', 'this', 'that', 'these', 'those', 'their', 'them', 'they', 'he', 'she',
                  'his', 'her', 'him', 'we', 'us', 'our', 'you', 'your', 'i', 'me', 'my', 'myself',
                  'yourself', 'himself', 'herself', 'itself', 'themselves', 'ourselves', 'yourselves',
                  'which', 'who', 'whom', 'whose', 'what', 'when', 'where', 'why', 'how', 'if', 'then',
                  'else', 'so', 'because', 'while', 'each', 'such', 'only', 'very', 'some', 'will',
                  'would', 'should', 'could', 'can', 'may', 'might', 'must', 'shall', 'not', 'no',
                  'nor', 'all', 'any', 'both', 'either', 'neither', 'few', 'many', 'more', 'most',
                  'other', 'some', 'such', 'than', 'too', 'very', 'just', 'ever', 'also'}
    
    # Convert to lowercase and tokenize
    anchor_words = {word.lower() for word in re.findall(r'\b\w+\b', anchor_text)}
    title_words = {word.lower() for word in re.findall(r'\b\w+\b', page_title)}
    
    # Remove stop words
    anchor_meaningful = anchor_words - stop_words
    title_meaningful = title_words - stop_words
    
    # Check for meaningful overlaps
    overlaps = anchor_meaningful.intersection(title_meaningful)
    
    return len(overlaps) > 0

def generate_article(keyword: str, semantic_structure: Dict, related_keywords: List[Dict], 
                     serp_features: List[Dict], paa_questions: List[Dict], openai_api_key: str) -> Tuple[str, bool]:
    """
    Generate comprehensive article with detailed paragraphs and natural language flow
    Returns: article_content, success_status
    """
    try:
        openai.api_key = openai_api_key
        
        # Ensure semantic_structure is valid
        if not semantic_structure:
            semantic_structure = {"h1": f"Guide to {keyword}", "sections": []}
        
        # Get default H1 if not present
        h1 = semantic_structure.get('h1', f"Complete Guide to {keyword}")
        
        # Prepare section structure with error handling
        sections_str = ""
        for section in semantic_structure.get('sections', []):
            if section and isinstance(section, dict) and 'h2' in section:
                sections_str += f"- {section.get('h2')}\n"
                for subsection in section.get('subsections', []):
                    if subsection and isinstance(subsection, dict) and 'h3' in subsection:
                        sections_str += f"  - {subsection.get('h3')}\n"
        
        # Add default section if none exist
        if not sections_str:
            sections_str = f"- Introduction to {keyword}\n- Key Benefits\n- How to Use\n- Tips and Best Practices\n- Conclusion\n"
        
        # Prepare related keywords with error handling
        related_kw_list = []
        if related_keywords and isinstance(related_keywords, list):
            for kw in related_keywords[:10]:
                if kw and isinstance(kw, dict) and 'keyword' in kw:
                    related_kw_list.append(kw.get('keyword', ''))
        
        # Add default keywords if none exist
        if not related_kw_list:
            related_kw_list = [f"{keyword} guide", f"best {keyword}", f"{keyword} tips", f"how to use {keyword}"]
        
        related_kw_str = ", ".join(related_kw_list)
        
        # Prepare SERP features with error handling
        serp_features_list = []
        if serp_features and isinstance(serp_features, list):
            for feature in serp_features[:5]:
                if feature and isinstance(feature, dict) and 'feature_type' in feature:
                    count = feature.get('count', 1)
                    serp_features_list.append(f"{feature.get('feature_type')} ({count})")
        
        # Add default features if none exist
        if not serp_features_list:
            serp_features_list = ["featured snippet", "people also ask", "images"]
        
        serp_features_str = ", ".join(serp_features_list)
        
        # Prepare People Also Asked questions
        paa_str = ""
        if paa_questions and isinstance(paa_questions, list):
            for i, question in enumerate(paa_questions[:5], 1):
                if question and isinstance(question, dict) and 'question' in question:
                    paa_str += f"{i}. {question.get('question', '')}\n"
        
        # Generate article with improved instructions for natural language flow
        response = openai.ChatCompletion.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": """You are an expert content writer with deep subject matter expertise.
                Write in a straightforward, authoritative style with no fluff or filler phrases.
                Focus on providing factual, specific information with concrete examples and evidence.

                Writing guidelines:
                1. Vary sentence length and structure - combine short sentences into longer, flowing ones where appropriate
                2. Avoid excessive repetition of the main topic terms - use pronouns and alternative phrasings 
                3. Create natural paragraph flow with proper transitions between ideas
                4. Use varied vocabulary and sentence constructions
                5. Write as a human expert would, with rhythm and cadence in your prose"""},
                {"role": "user", "content": f"""
                Write a comprehensive, expert-level article about "{keyword}".
                
                Use this semantic structure:
                H1: {h1}
                
                Sections:
                {sections_str}
                
                Content requirements:
                1. Write thorough, information-rich paragraphs (minimum 150 words per section)
                2. Include specific examples, data points, and practical details in each section
                3. Use direct, precise language - STRICTLY AVOID:
                   - Filler words like "Additionally", "Moreover", "Furthermore", "Also"
                   - Phrases like "For example", "For instance", "Similarly", "In particular"
                   - Transitions like "It's worth noting", "It's important to understand"
                4. Start sentences with the main point - not with filler transitions
                5. Keep headings concise and descriptive (max 8 words)
                6. Maintain a professional tone throughout
                7. Include the related keywords naturally: {related_kw_str}
                8. Answer these 'People Also Asked' questions within the article text:
                {paa_str}
                9. Optimize for these SERP features: {serp_features_str}
                10. IMPORTANT: After writing the first draft, revise to:
                    - Combine short, choppy sentences into longer, flowing ones
                    - Vary your references to the main topic (don't repeat "{keyword}" excessively)
                    - Use pronouns, descriptive alternatives, and varied terms to reference the topic
                    - Ensure natural transitions between ideas without using filler phrases
                
                Format the article with proper HTML:
                - Main title in <h1> tags
                - Section headings in <h2> tags
                - Subsection headings in <h3> tags
                - Paragraphs in <p> tags
                - Use <ul>, <li> for bullet points and <ol>, <li> for numbered lists
                
                The article should be authoritative, factually accurate, and comprehensive.
                Aim for around 1,800-2,200 words total with substantive, detailed content in each section.
                """}
            ],
            temperature=0.5
        )
        
        article_content = response.choices[0].message.content
        return article_content, True
    
    except Exception as e:
        error_msg = f"Exception in generate_article: {str(e)}"
        logger.error(error_msg)
        return "", False

###############################################################################
# 8. Internal Linking
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

def generate_internal_links_with_embeddings(article_content: str, pages_with_embeddings: List[Dict], 
                                          openai_api_key: str, word_count: int) -> Tuple[str, List[Dict], bool]:
    """
    Generate internal links by replacing existing text with linked versions
    Returns: article_with_links, links_added, success_status
    """
    try:
        openai.api_key = openai_api_key
        
        # Calculate max links based on word count
        max_links = min(10, max(2, int(word_count / 1000) * 8))  # Ensure at least 2 links
        
        # Log for debugging
        logger.info(f"Generating up to {max_links} internal links for content with {word_count} words")
        
        # Convert pages to simple format for prompt
        pages_str = "\n".join([f"URL: {p['url']}, Title: {p['title']}, Description: {p['description']}" 
                             for p in pages_with_embeddings[:30]])
        
        # Use a different approach - two-step process
        # First, generate just the links to add
        response = openai.ChatCompletion.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": "You are an SEO expert who identifies perfect places for internal links."},
                {"role": "user", "content": f"""
                I need you to identify {max_links} opportunities to add internal links to this HTML article.

                For each link opportunity:
                1. Select a 2-6 word phrase from the EXISTING article text
                2. Match it with the most relevant page from the list below
                3. Each URL should only be used ONCE
                4. Choose phrases that are natural anchor text
                5. CRITICAL: The anchor text MUST contain at least one meaningful word that appears in the title of the linked page
                6. Skip common/stop words when considering matches (words like "the", "of", "and", etc.)
                7. Only suggest links where there's a clear topical match

                Available pages:
                {pages_str}

                HTML article:
                {article_content}

                Return JUST A JSON ARRAY of your link suggestions:
                [
                  {{
                    "url": "page URL to link to",
                    "anchor_text": "exact text from article to turn into a link",
                    "surrounding_text": "...text before [anchor text] text after..."
                  }},
                  ...more link suggestions...
                ]
                """
                }
            ],
            response_format={"type": "json_object"},  # Force JSON response
            temperature=0.4
        )
        
        # Parse the response
        try:
            link_suggestions = json.loads(response.choices[0].message.content)
            
            # If we have suggestions key, use that, otherwise use the whole object
            if "suggestions" in link_suggestions:
                suggestions = link_suggestions["suggestions"]
            else:
                # Find the first array in the JSON object
                for key, value in link_suggestions.items():
                    if isinstance(value, list) and len(value) > 0:
                        suggestions = value
                        break
                else:
                    suggestions = []
            
            logger.info(f"Generated {len(suggestions)} link suggestions")
            
            # Filter suggestions to ensure they have semantic matches
            verified_suggestions = []
            for suggestion in suggestions:
                anchor_text = suggestion.get("anchor_text", "")
                url = suggestion.get("url", "")
                
                # Find the page title
                page_title = ""
                for page in pages_with_embeddings:
                    if page.get('url') == url:
                        page_title = page.get('title', "")
                        break
                
                # Verify semantic match
                if verify_semantic_match(anchor_text, page_title):
                    verified_suggestions.append(suggestion)
                    logger.info(f"Verified match: '{anchor_text}' matches with '{page_title}'")
                else:
                    logger.warning(f"Rejected match: '{anchor_text}' has no semantic match with '{page_title}'")
            
            logger.info(f"After verification: {len(verified_suggestions)} of {len(suggestions)} suggestions approved")
            
            # Apply links using the more precise function
            article_with_links = apply_internal_links(article_content, verified_suggestions)
            
            # Format the final list of links
            links_added = []
            for suggestion in verified_suggestions:
                if suggestion.get("url") and suggestion.get("anchor_text"):
                    links_added.append({
                        "url": suggestion.get("url", ""),
                        "anchor_text": suggestion.get("anchor_text", ""),
                        "context": suggestion.get("surrounding_text", suggestion.get("anchor_text", ""))
                    })
            
            # Check if any links were added
            if article_content == article_with_links:
                logger.warning("No links were added to the article")
                return article_content, [], False
            
            return article_with_links, links_added, True
            
        except json.JSONDecodeError:
            logger.error("Failed to parse JSON response for link suggestions")
            return article_content, [], False
            
    except Exception as e:
        error_msg = f"Exception in generate_internal_links_with_embeddings: {str(e)}"
        logger.error(error_msg)
        logger.error(traceback.format_exc())
        return article_content, [], False

def apply_internal_links(article_content: str, link_suggestions: List[Dict]) -> str:
    """
    Apply internal links to article content with improved precision
    """
    # Sort suggestions by anchor text length (descending) to avoid substring issues
    link_suggestions = sorted(link_suggestions, key=lambda x: len(x.get('anchor_text', '')), reverse=True)
    
    # Create HTML soup to preserve structure
    soup = BeautifulSoup(article_content, 'html.parser')
    
    # Track paragraphs we've already processed to avoid duplicate replacements
    processed_tags = set()
    replaced_anchors = set()
    
    # Process each link suggestion
    for suggestion in link_suggestions:
        url = suggestion.get('url', '')
        anchor_text = suggestion.get('anchor_text', '')
        
        if not url or not anchor_text or anchor_text in replaced_anchors:
            continue
            
        # Find all text elements
        for tag in soup.find_all(['p', 'li', 'h2', 'h3']):
            # Skip if we've already processed this tag
            tag_id = id(tag)
            if tag_id in processed_tags:
                continue
                
            # Skip headings that contain the exact anchor text (avoid changing heading content)
            if tag.name in ['h1', 'h2', 'h3'] and tag.get_text().strip() == anchor_text:
                continue
                
            # Check if anchor text is in this tag
            tag_text = str(tag)
            if anchor_text in tag_text and '<a ' not in tag_text:
                # Replace text with linked version but preserve HTML
                linked_html = tag_text.replace(
                    anchor_text, 
                    f'<a href="{url}">{anchor_text}</a>',
                    1  # Only replace first occurrence
                )
                
                # Replace the tag with updated HTML
                new_tag = BeautifulSoup(linked_html, 'html.parser')
                tag.replace_with(new_tag)
                
                # Mark this anchor as replaced and tag as processed
                replaced_anchors.add(anchor_text)
                processed_tags.add(tag_id)
                break
    
    # Return the updated HTML
    return str(soup)

###############################################################################
# 9. Document Generation
###############################################################################

def create_word_document(keyword: str, serp_results: List[Dict], related_keywords: List[Dict],
                        semantic_structure: Dict, article_content: str, meta_title: str, 
                        meta_description: str, paa_questions: List[Dict],
                        internal_links: List[Dict] = None) -> Tuple[BytesIO, bool]:
    """
    Create Word document with all components
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
        
        # Section 3: Semantic Structure
        doc.add_heading('Recommended Content Structure', level=1)
        
        doc.add_paragraph(f"Recommended H1: {semantic_structure.get('h1', '')}")
        
        for i, section in enumerate(semantic_structure.get('sections', []), 1):
            doc.add_paragraph(f"H2 Section {i}: {section.get('h2', '')}")
            
            for j, subsection in enumerate(section.get('subsections', []), 1):
                doc.add_paragraph(f"    H3 Subsection {j}: {subsection.get('h3', '')}")
        
        # Section 4: Generated Article
        doc.add_heading('Generated Article with Internal Links', level=1)
        
        # Parse HTML content and add to document, preserving links
        soup = BeautifulSoup(article_content, 'html.parser')
        
        for element in soup.find_all(['h1', 'h2', 'h3', 'p', 'ul', 'ol']):
            if element.name in ['h1', 'h2', 'h3']:
                level = int(element.name[1])
                doc.add_heading(element.get_text(), level=level)
            elif element.name == 'p':
                p = doc.add_paragraph()
                
                # Process paragraph content including links
                for content in element.contents:
                    if isinstance(content, str):
                        # Plain text
                        p.add_run(content)
                    elif content.name == 'a':
                        # Add link as hyperlinked text with blue color
                        url = content.get('href', '')
                        text = content.get_text()
                        if url and text:
                            run = p.add_run(text)
                            run.font.color.rgb = RGBColor(0, 0, 255)  # Blue
                            run.underline = True
                            # Note: This creates the visual appearance but not functional hyperlinks
                    else:
                        # Other tags
                        p.add_run(content.get_text())
                        
            elif element.name == 'ul':
                for li in element.find_all('li'):
                    p = doc.add_paragraph(style='List Bullet')
                    for content in li.contents:
                        if isinstance(content, str):
                            p.add_run(content)
                        elif content.name == 'a':
                            url = content.get('href', '')
                            text = content.get_text()
                            if url and text:
                                run = p.add_run(text)
                                run.font.color.rgb = RGBColor(0, 0, 255)
                                run.underline = True
                        else:
                            p.add_run(content.get_text())
                            
            elif element.name == 'ol':
                for i, li in enumerate(element.find_all('li')):
                    p = doc.add_paragraph(style='List Number')
                    for content in li.contents:
                        if isinstance(content, str):
                            p.add_run(content)
                        elif content.name == 'a':
                            url = content.get('href', '')
                            text = content.get_text()
                            if url and text:
                                run = p.add_run(text)
                                run.font.color.rgb = RGBColor(0, 0, 255)
                                run.underline = True
                        else:
                            p.add_run(content.get_text())
        
        # Section 5: Internal Linking (if provided)
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
# 10. Main Streamlit App
###############################################################################

def main():
    st.title("📊 SEO Analysis & Content Generator")
    
    # Sidebar for API credentials
    st.sidebar.header("API Credentials")
    
    dataforseo_login = st.sidebar.text_input("DataForSEO API Login", type="password")
    dataforseo_password = st.sidebar.text_input("DataForSEO API Password", type="password")
    
    openai_api_key = st.sidebar.text_input("OpenAI API Key", type="password")
    
    # Initialize session state
    if 'results' not in st.session_state:
        st.session_state.results = {}
    
    # Main content
    tabs = st.tabs([
        "Input & SERP Analysis", 
        "Content Analysis", 
        "Article Generation", 
        "Internal Linking", 
        "SEO Brief"
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
            elif not openai_api_key:
                st.error("Please enter OpenAI API key")
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
                if not openai_api_key:
                    st.error("Please enter OpenAI API key")
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
                        
                        # Analyze semantic structure
                        st.text("Analyzing semantic structure...")
                        semantic_structure, structure_success = analyze_semantic_structure(
                            scraped_contents, openai_api_key
                        )
                        
                        if structure_success:
                            st.session_state.results['semantic_structure'] = semantic_structure
                            
                            st.subheader("Recommended Semantic Structure")
                            st.write(f"**H1:** {semantic_structure.get('h1', '')}")
                            
                            for i, section in enumerate(semantic_structure.get('sections', []), 1):
                                st.write(f"**H2 {i}:** {section.get('h2', '')}")
                                
                                for j, subsection in enumerate(section.get('subsections', []), 1):
                                    st.write(f"  - **H3 {i}.{j}:** {subsection.get('h3', '')}")
                            
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
    
    # Tab 3: Article Generation
    with tabs[2]:
        st.header("Article Generation")
        
        if 'semantic_structure' not in st.session_state.results:
            st.warning("Please complete content analysis first (in the 'Content Analysis' tab)")
        else:
            if st.button("Generate Article and Meta Tags"):
                if not openai_api_key:
                    st.error("Please enter OpenAI API key")
                else:
                    with st.spinner("Generating article and meta tags..."):
                        start_time = time.time()
                        
                        # Generate article (updated to include PAA questions)
                        article_content, article_success = generate_article(
                            st.session_state.results['keyword'],
                            st.session_state.results['semantic_structure'],
                            st.session_state.results.get('related_keywords', []),
                            st.session_state.results.get('serp_features', []),
                            st.session_state.results.get('paa_questions', []),
                            openai_api_key
                        )
                        
                        if article_success and article_content:
                            st.session_state.results['article_content'] = article_content
                            
                            # Generate meta title and description
                            meta_title, meta_description, meta_success = generate_meta_tags(
                                st.session_state.results['keyword'],
                                st.session_state.results['semantic_structure'],
                                st.session_state.results.get('related_keywords', []),
                                openai_api_key
                            )
                            
                            if meta_success:
                                st.session_state.results['meta_title'] = meta_title
                                st.session_state.results['meta_description'] = meta_description
                                
                                st.subheader("Meta Tags")
                                st.write(f"**Meta Title:** {meta_title}")
                                st.write(f"**Meta Description:** {meta_description}")
                            
                            st.subheader("Generated Article")
                            st.markdown(article_content, unsafe_allow_html=True)
                            
                            st.success(f"Content generation completed in {format_time(time.time() - start_time)}")
                        else:
                            st.error("Failed to generate article. Please try again.")
            
            # Show previously generated article if available
            if 'article_content' in st.session_state.results:
                if 'meta_title' in st.session_state.results:
                    st.subheader("Previously Generated Meta Tags")
                    st.write(f"**Meta Title:** {st.session_state.results['meta_title']}")
                    st.write(f"**Meta Description:** {st.session_state.results['meta_description']}")
                
                st.subheader("Previously Generated Article")
                st.markdown(st.session_state.results['article_content'], unsafe_allow_html=True)
    
    # Tab 4: Internal Linking
    with tabs[3]:
        st.header("Internal Linking")
        
        if 'article_content' not in st.session_state.results:
            st.warning("Please generate an article first (in the 'Article Generation' tab)")
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
                    st.error("Please enter OpenAI API key")
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
                            
                            # Generate embeddings for site pages
                            pages_with_embeddings, embed_success = embed_site_pages(
                                pages, openai_api_key, batch_size
                            )
                            
                            if embed_success:
                                # Count words in the article
                                article_content = st.session_state.results['article_content']
                                word_count = len(re.findall(r'\w+', article_content))
                                
                                status_text.text(f"Analyzing article content and generating internal links...")
                                
                                # Generate internal links (updated function)
                                article_with_links, links_added, links_success = generate_internal_links_with_embeddings(
                                    article_content, pages_with_embeddings, openai_api_key, word_count
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
        
        if 'article_content' not in st.session_state.results:
            st.warning("Please generate an article first (in the 'Article Generation' tab)")
        else:
            if st.button("Generate SEO Brief"):
                with st.spinner("Generating SEO brief..."):
                    start_time = time.time()
                    
                    # Use article with internal links if available, otherwise use regular article
                    article_content = st.session_state.results.get('article_with_links', 
                                                                 st.session_state.results['article_content'])
                    
                    internal_links = st.session_state.results.get('internal_links', None)
                    
                    # Get meta title and description
                    meta_title = st.session_state.results.get('meta_title', 
                                                             f"{st.session_state.results['keyword']} - Complete Guide")
                    
                    meta_description = st.session_state.results.get('meta_description', 
                                                                  f"Learn everything about {st.session_state.results['keyword']} in our comprehensive guide.")
                    
                    # Get PAA questions
                    paa_questions = st.session_state.results.get('paa_questions', [])
                    
                    # Create Word document (updated to include PAA questions)
                    doc_stream, doc_success = create_word_document(
                        st.session_state.results['keyword'],
                        st.session_state.results['organic_results'],
                        st.session_state.results.get('related_keywords', []),
                        st.session_state.results['semantic_structure'],
                        article_content,
                        meta_title,
                        meta_description,
                        paa_questions,
                        internal_links
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
                ("Semantic Structure", 'semantic_structure' in st.session_state.results),
                ("Meta Title & Description", 'meta_title' in st.session_state.results),
                ("Generated Article", 'article_content' in st.session_state.results),
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

if __name__ == "__main__":
    main()
