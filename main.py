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
    page_title="SEO Content Brief Generator",
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
# 5. Content Analysis with GPT-4o-mini
###############################################################################

def extract_competitor_insights(competitor_contents: List[Dict], openai_api_key: str, keyword: str) -> Tuple[Dict, bool]:
    """
    Extract core themes, headings, and content insights from competitor pages
    specifically for SEO brief creation. This is a focused extraction that 
    helps with quick SEO brief ideation.
    
    Returns: competitor_insights, success_status
    """
    try:
        # Set OpenAI API key
        client = openai.OpenAI(api_key=openai_api_key)
        
        # Extract all headings from competitor content
        all_headings = {"h1": [], "h2": [], "h3": []}
        for content in competitor_contents:
            headings = content.get('headings', {})
            if headings:
                for level in ['h1', 'h2', 'h3']:
                    if level in headings:
                        all_headings[level].extend(headings[level])
        
        # Combine all content for holistic analysis
        combined_content = ""
        for i, c in enumerate(competitor_contents):
            title = c.get('title', f'Competitor {i+1}')
            url = c.get('url', '')
            content = c.get('content', '')
            
            # Add content with source identification
            combined_content += f"SOURCE: {title} ({url})\n\n{content}\n\n" + "-" * 40 + "\n\n"
        
        # Prepare summarized content if it's too long
        if len(combined_content) > 20000:
            combined_content = combined_content[:20000]
        
        # Use GPT-4o-mini to extract insights from competitor content
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            max_tokens=1500,
            messages=[
                {"role": "system", "content": "You are an SEO expert who extracts valuable insights from competitor content to create effective content briefs."},
                {"role": "user", "content": f"""
                Analyze the following content from top-ranking pages for the keyword "{keyword}" and extract the most important insights for creating a comprehensive content brief.
                
                Focus ONLY on extracting these elements (no analysis or recommendations):
                
                1. Core themes (main topics that appear across multiple competitors)
                2. Common headings structure used by competitors
                3. Key points/facts consistently mentioned across competitors
                4. Types of content included (lists, FAQs, stats, examples, etc.)
                5. Important terminology and definitions
                
                Format your response as JSON:
                {{
                    "core_themes": [
                        {{
                            "theme": "Main theme name",
                            "importance": "High/Medium/Low",
                            "description": "Brief description of what this theme covers"
                        }}
                    ],
                    "common_heading_structure": [
                        {{
                            "section_type": "Introduction/Main Section/Conclusion etc.",
                            "example_headings": ["Heading example 1", "Heading example 2"],
                            "typical_content": "Brief description of what's typically included in this section"
                        }}
                    ],
                    "key_points": [
                        {{
                            "point": "Specific point/fact",
                            "frequency": "Mentioned in X out of Y competitors",
                            "context": "When/where this point is typically mentioned"
                        }}
                    ],
                    "content_types": [
                        {{
                            "type": "Lists/FAQs/Stats/Examples/etc.",
                            "usage": "How competitors typically use this content type",
                            "examples": ["Brief example 1", "Brief example 2"]
                        }}
                    ],
                    "terminology": [
                        {{
                            "term": "Specific term",
                            "definition": "How it's defined",
                            "importance": "High/Medium/Low"
                        }}
                    ]
                }}
                
                Be concise and specific. Focus on extracting factual information, not making recommendations.
                
                Content to analyze:
                {combined_content}
                """}
            ],
            temperature=0.3
        )
        
        # Extract and parse JSON response
        content = response.choices[0].message.content
        json_match = re.search(r'({.*})', content, re.DOTALL)
        if json_match:
            content = json_match.group(1)
        
        competitor_insights = json.loads(content)
        
        # Add extracted headings data to the insights
        competitor_insights['extracted_headings'] = all_headings
        
        return competitor_insights, True
    
    except Exception as e:
        error_msg = f"Exception in extract_competitor_insights: {str(e)}"
        logger.error(error_msg)
        return {}, False

def generate_content_brief(keyword: str, competitor_insights: Dict, paa_questions: List[Dict], 
                          related_keywords: List[Dict], openai_api_key: str) -> Tuple[str, bool]:
    """
    Generate a concise, focused content brief based on competitor insights
    Returns: brief_content, success_status
    """
    try:
        # Set OpenAI API key
        client = openai.OpenAI(api_key=openai_api_key)
        
        # Format competitor insights for the prompt
        core_themes = json.dumps(competitor_insights.get('core_themes', []), indent=2)
        heading_structure = json.dumps(competitor_insights.get('common_heading_structure', []), indent=2)
        key_points = json.dumps(competitor_insights.get('key_points', []), indent=2)
        content_types = json.dumps(competitor_insights.get('content_types', []), indent=2)
        terminology = json.dumps(competitor_insights.get('terminology', []), indent=2)
        
        # Format extracted headings
        extracted_headings = competitor_insights.get('extracted_headings', {})
        h1_headings = json.dumps(extracted_headings.get('h1', []), indent=2)
        h2_headings = json.dumps(extracted_headings.get('h2', []), indent=2)
        h3_headings = json.dumps(extracted_headings.get('h3', []), indent=2)
        
        # Format PAA questions
        paa_questions_text = ""
        if paa_questions:
            paa_questions_text = "People Also Asked Questions:\n"
            for i, q in enumerate(paa_questions, 1):
                paa_questions_text += f"{i}. {q.get('question', '')}\n"
        
        # Format related keywords
        related_keywords_text = ", ".join([kw.get('keyword', '') for kw in related_keywords[:10] if kw.get('keyword')])
        
        # Generate content brief
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            max_tokens=1800,
            messages=[
                {"role": "system", "content": "You are an expert content strategist who creates concise, actionable content briefs based on competitor analysis."},
                {"role": "user", "content": f"""
                Create a concise, actionable content brief for the keyword: "{keyword}"
                
                Use these competitor insights to inform the brief:
                
                CORE THEMES:
                {core_themes}
                
                COMMON HEADING STRUCTURE:
                {heading_structure}
                
                KEY POINTS CONSISTENTLY MENTIONED:
                {key_points}
                
                CONTENT TYPES USED BY COMPETITORS:
                {content_types}
                
                IMPORTANT TERMINOLOGY:
                {terminology}
                
                COMPETITOR H1 HEADINGS:
                {h1_headings}
                
                COMPETITOR H2 HEADINGS:
                {h2_headings}
                
                COMPETITOR H3 HEADINGS:
                {h3_headings}
                
                PEOPLE ALSO ASKED QUESTIONS:
                {paa_questions_text}
                
                RELATED KEYWORDS:
                {related_keywords_text}
                
                Create a content brief that includes:
                
                1. Suggested article title (H1)
                2. Recommended main sections (H2s) with brief descriptions of what to include
                3. Suggested subsections (H3s) where appropriate
                4. Key points that must be covered
                5. Questions that should be answered
                6. Recommended content types (lists, tables, examples, etc.)
                7. Important terminology to include
                8. Target word count for each section
                
                FORMAT:
                - Use Markdown formatting with appropriate headings
                - Be concise and actionable
                - Focus on creating an outline that would be easy for a writer to follow
                - Include specific guidance on what makes competitors successful
                
                The brief should be CONCISE and PRACTICAL - avoid any theoretical explanations or fluff. 
                Focus on specific, actionable guidance based directly on what's working for competitors.
                """}
            ],
            temperature=0.4
        )
        
        # Extract brief content
        brief_content = response.choices[0].message.content
        
        return brief_content, True
    
    except Exception as e:
        error_msg = f"Exception in generate_content_brief: {str(e)}"
        logger.error(error_msg)
        return "", False

def generate_meta_tags(keyword: str, semantic_structure: Dict, related_keywords: List[Dict], term_data: Dict, 
                      openai_api_key: str) -> Tuple[str, str, bool]:
    """
    Generate optimized meta title and description for the content
    Returns: meta_title, meta_description, success_status
    """
    try:
        # Set OpenAI API key
        client = openai.OpenAI(api_key=openai_api_key)
        
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
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            max_tokens=300,
            messages=[
                {"role": "system", "content": "You are an SEO specialist who creates optimized meta tags."},
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
# 6. Main Streamlit App
###############################################################################

def main():
    st.title("🔍 SEO Content Brief Generator")
    
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
        "Competitor Content Analysis", 
        "Content Brief Generation"
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
    
    # Tab 2: Competitor Content Analysis
    with tabs[1]:
        st.header("Competitor Content Analysis")
        
        if 'organic_results' not in st.session_state.results:
            st.warning("Please fetch SERP data first (in the 'Input & SERP Analysis' tab)")
        else:
            if st.button("Analyze Competitor Content"):
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
                                # Extract headings from the URL
                                headings = extract_headings(result['url'])
                                
                                scraped_contents.append({
                                    'url': result['url'],
                                    'title': result['title'],
                                    'content': content,
                                    'headings': headings
                                })
                            else:
                                st.warning(f"Could not scrape content from {result['url']}")
                            
                            progress_bar.progress((i + 1) / len(st.session_state.results['organic_results']))
                        
                        if not scraped_contents:
                            st.error("Could not scrape content from any URLs. Please try a different keyword.")
                            return
                            
                        st.session_state.results['scraped_contents'] = scraped_contents
                        
                        # Extract insights from competitor content
                        st.text("Extracting insights from competitor content...")
                        competitor_insights, insights_success = extract_competitor_insights(
                            scraped_contents, 
                            openai_api_key, 
                            st.session_state.results['keyword']
                        )
                        
                        if insights_success:
                            st.session_state.results['competitor_insights'] = competitor_insights
                            
                            # Display core themes
                            st.subheader("Core Themes")
                            if 'core_themes' in competitor_insights:
                                themes_df = pd.DataFrame(competitor_insights['core_themes'])
                                st.dataframe(themes_df)
                            
                            # Display common heading structure
                            st.subheader("Common Heading Structure")
                            if 'common_heading_structure' in competitor_insights:
                                for section in competitor_insights['common_heading_structure']:
                                    st.markdown(f"**{section.get('section_type', '')}**")
                                    st.markdown(f"*Example headings:* {', '.join(section.get('example_headings', []))}")
                                    st.markdown(f"*Typical content:* {section.get('typical_content', '')}")
                                    st.markdown("---")
                            
                            # Display extracted headings
                            with st.expander("View All Extracted Headings"):
                                if 'extracted_headings' in competitor_insights:
                                    extracted = competitor_insights['extracted_headings']
                                    
                                    st.markdown("### H1 Headings")
                                    for h in extracted.get('h1', []):
                                        st.markdown(f"- {h}")
                                    
                                    st.markdown("### H2 Headings")
                                    for h in extracted.get('h2', []):
                                        st.markdown(f"- {h}")
                                    
                                    st.markdown("### H3 Headings")
                                    for h in extracted.get('h3', []):
                                        st.markdown(f"- {h}")
                            
                            # Display key points
                            st.subheader("Key Points")
                            if 'key_points' in competitor_insights:
                                points_df = pd.DataFrame(competitor_insights['key_points'])
                                st.dataframe(points_df)
                            
                            # Display content types
                            st.subheader("Content Types Used by Competitors")
                            if 'content_types' in competitor_insights:
                                for content_type in competitor_insights['content_types']:
                                    st.markdown(f"**{content_type.get('type', '')}**")
                                    st.markdown(f"*Usage:* {content_type.get('usage', '')}")
                                    st.markdown(f"*Examples:* {', '.join(content_type.get('examples', []))}")
                                    st.markdown("---")
                            
                            # Display terminology
                            st.subheader("Important Terminology")
                            if 'terminology' in competitor_insights:
                                terms_df = pd.DataFrame(competitor_insights['terminology'])
                                st.dataframe(terms_df)
                            
                            st.success(f"Competitor analysis completed in {format_time(time.time() - start_time)}")
                        else:
                            st.error("Failed to extract insights from competitor content")
            
            # Show previously analyzed data if available
            if 'competitor_insights' in st.session_state.results:
                insights = st.session_state.results['competitor_insights']
                
                st.subheader("Previously Analyzed Competitor Insights")
                
                # Display core themes
                if 'core_themes' in insights:
                    with st.expander("Core Themes"):
                        themes_df = pd.DataFrame(insights['core_themes'])
                        st.dataframe(themes_df)
                
                # Display common heading structure
                if 'common_heading_structure' in insights:
                    with st.expander("Common Heading Structure"):
                        for section in insights['common_heading_structure']:
                            st.markdown(f"**{section.get('section_type', '')}**")
                            st.markdown(f"*Example headings:* {', '.join(section.get('example_headings', []))}")
                            st.markdown(f"*Typical content:* {section.get('typical_content', '')}")
                            st.markdown("---")
                
                # Display key points
                if 'key_points' in insights:
                    with st.expander("Key Points"):
                        points_df = pd.DataFrame(insights['key_points'])
                        st.dataframe(points_df)
                
                # Display content types
                if 'content_types' in insights:
                    with st.expander("Content Types Used by Competitors"):
                        for content_type in insights['content_types']:
                            st.markdown(f"**{content_type.get('type', '')}**")
                            st.markdown(f"*Usage:* {content_type.get('usage', '')}")
                            st.markdown(f"*Examples:* {', '.join(content_type.get('examples', []))}")
                            st.markdown("---")
                
                # Display terminology
                if 'terminology' in insights:
                    with st.expander("Important Terminology"):
                        terms_df = pd.DataFrame(insights['terminology'])
                        st.dataframe(terms_df)
    
    # Tab 3: Content Brief Generation
    with tabs[2]:
        st.header("Content Brief Generation")
        
        if 'competitor_insights' not in st.session_state.results:
            st.warning("Please analyze competitor content first (in the 'Competitor Content Analysis' tab)")
        else:
            if st.button("Generate Content Brief"):
                if not openai_api_key:
                    st.error("Please enter OpenAI API key")
                else:
                    with st.spinner("Generating content brief based on competitor insights..."):
                        start_time = time.time()
                        
                        # Generate content brief
                        brief_content, brief_success = generate_content_brief(
                            st.session_state.results['keyword'],
                            st.session_state.results['competitor_insights'],
                            st.session_state.results.get('paa_questions', []),
                            st.session_state.results.get('related_keywords', []),
                            openai_api_key
                        )
                        
                        if brief_success and brief_content:
                            st.session_state.results['brief_content'] = brief_content
                            
                            # Display the brief content
                            st.subheader(f"Content Brief for: {st.session_state.results['keyword']}")
                            st.markdown(brief_content)
                            
                            # Generate meta tags if needed
                            if 'meta_title' not in st.session_state.results or 'meta_description' not in st.session_state.results:
                                # Extract semantic structure from competitor insights
                                semantic_structure = {
                                    'h1': '',
                                    'sections': []
                                }
                                
                                if 'common_heading_structure' in st.session_state.results['competitor_insights']:
                                    structure = st.session_state.results['competitor_insights']['common_heading_structure']
                                    for section in structure:
                                        if section.get('section_type') == 'Introduction':
                                            if section.get('example_headings') and len(section.get('example_headings', [])) > 0:
                                                semantic_structure['h1'] = section.get('example_headings')[0]
                                
                                meta_title, meta_description, meta_success = generate_meta_tags(
                                    st.session_state.results['keyword'],
                                    semantic_structure,
                                    st.session_state.results.get('related_keywords', []),
                                    {},  # No term data needed
                                    openai_api_key
                                )
                                
                                if meta_success:
                                    st.session_state.results['meta_title'] = meta_title
                                    st.session_state.results['meta_description'] = meta_description
                            
                            # Create Word document
                            doc = Document()
                            
                            # Add document title
                            doc.add_heading(f'Content Brief: {st.session_state.results["keyword"]}', 0)
                            
                            # Add date
                            doc.add_paragraph(f"Generated on: {time.strftime('%Y-%m-%d %H:%M:%S')}")
                            
                            # Add meta tags if available
                            if 'meta_title' in st.session_state.results and 'meta_description' in st.session_state.results:
                                doc.add_heading('Meta Tags', level=1)
                                meta_title_para = doc.add_paragraph()
                                meta_title_para.add_run("Meta Title: ").bold = True
                                meta_title_para.add_run(st.session_state.results['meta_title'])
                                
                                meta_desc_para = doc.add_paragraph()
                                meta_desc_para.add_run("Meta Description: ").bold = True
                                meta_desc_para.add_run(st.session_state.results['meta_description'])
                            
                            # Process brief content with markdown parsing
                            # Simple markdown to docx conversion
                            lines = brief_content.split('\n')
                            
                            for line in lines:
                                # Skip empty lines
                                if not line.strip():
                                    continue
                                
                                # Check for headings
                                if line.startswith('# '):
                                    doc.add_heading(line[2:], level=1)
                                elif line.startswith('## '):
                                    doc.add_heading(line[3:], level=2)
                                elif line.startswith('### '):
                                    doc.add_heading(line[4:], level=3)
                                
                                # Check for list items
                                elif line.startswith('- ') or line.startswith('* '):
                                    doc.add_paragraph(line[2:], style='List Bullet')
                                elif re.match(r'^\d+\.\s', line):
                                    text = re.sub(r'^\d+\.\s', '', line)
                                    doc.add_paragraph(text, style='List Number')
                                
                                # Regular paragraph
                                else:
                                    doc.add_paragraph(line)
                            
                            # Save document to BytesIO stream
                            doc_stream = BytesIO()
                            doc.save(doc_stream)
                            doc_stream.seek(0)
                            
                            # Store in session state
                            st.session_state.results['brief_doc'] = doc_stream
                            
                            # Create download button
                            st.download_button(
                                label="Download Content Brief",
                                data=doc_stream,
                                file_name=f"content_brief_{st.session_state.results['keyword'].replace(' ', '_')}.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                            )
                            
                            # Show meta tags if available
                            if 'meta_title' in st.session_state.results and 'meta_description' in st.session_state.results:
                                st.subheader("Generated Meta Tags")
                                st.write(f"**Meta Title:** {st.session_state.results['meta_title']}")
                                st.write(f"**Meta Description:** {st.session_state.results['meta_description']}")
                            
                            st.success(f"Content brief generation completed in {format_time(time.time() - start_time)}")
                        else:
                            st.error("Failed to generate content brief. Please try again.")
            
            # Show previously generated brief if available
            if 'brief_content' in st.session_state.results:
                st.subheader("Previously Generated Content Brief")
                
                # Show meta tags if available
                if 'meta_title' in st.session_state.results and 'meta_description' in st.session_state.results:
                    st.write(f"**Meta Title:** {st.session_state.results['meta_title']}")
                    st.write(f"**Meta Description:** {st.session_state.results['meta_description']}")
                
                # Show brief content
                st.markdown(st.session_state.results['brief_content'])
                
                # Show download button if available
                if 'brief_doc' in st.session_state.results:
                    st.download_button(
                        label="Download Content Brief",
                        data=st.session_state.results['brief_doc'],
                        file_name=f"content_brief_{st.session_state.results['keyword'].replace(' ', '_')}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key="download_previous_brief"  # Use unique key to avoid duplicates
                    )

if __name__ == "__main__":
    main()
