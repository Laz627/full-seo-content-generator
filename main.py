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

def fetch_serp_results(keyword: str, api_login: str, api_password: str) -> Tuple[List[Dict], List[Dict], bool]:
    """
    Fetch SERP results from DataForSEO API and classify pages
    Returns: organic_results, serp_features, success_status
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
                
                return organic_results, serp_features, True
            else:
                error_msg = f"API Error: {data.get('status_message')}"
                logger.error(error_msg)
                return [], [], False
        else:
            error_msg = f"HTTP Error: {response.status_code} - {response.text}"
            logger.error(error_msg)
            return [], [], False
    
    except Exception as e:
        error_msg = f"Exception in fetch_serp_results: {str(e)}"
        logger.error(error_msg)
        logger.error(traceback.format_exc())
        return [], [], False

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

def fetch_related_keywords_dataforseo(keyword: str, api_login: str, api_password: str) -> Tuple[List[Dict], bool]:
    """
    Fetch related keywords from DataForSEO Keyword Suggestions API
    Updated to match the API response structure in the sample
    Returns: related_keywords, success_status
    """
    try:
        # Use the keyword_suggestions endpoint as shown in the sample
        url = "https://api.dataforseo.com/v3/dataforseo_labs/google/keyword_suggestions/live"
        headers = {
            'Content-Type': 'application/json',
        }
        
        # Prepare request data following the sample format
        post_data = [{
            "keyword": keyword,
            "location_code": 2840,  # USA
            "language_code": "en",
            "include_serp_info": True,
            "include_seed_keyword": True,
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
            
            # Log complete response for debugging
            logger.info(f"API Response: {json.dumps(data, indent=2)}")
            
            # Basic validation of response
            if data.get('status_code') != 20000:
                error_msg = f"API Error: {data.get('status_code')} - {data.get('status_message')}"
                logger.error(error_msg)
                return create_default_keywords(keyword), False
            
            # Check if we have tasks and results
            if not data.get('tasks') or len(data['tasks']) == 0 or not data['tasks'][0].get('result'):
                logger.warning(f"No tasks or results in API response")
                return create_default_keywords(keyword), False
            
            # Process according to the JSON structure in the sample
            result = data['tasks'][0]['result']
            related_keywords = []
            
            # First, try to extract the seed keyword data
            if len(result) > 0 and 'seed_keyword_data' in result[0]:
                seed_data = result[0]['seed_keyword_data']
                if 'keyword_info' in seed_data:
                    keyword_info = seed_data['keyword_info']
                    related_keywords.append({
                        'keyword': result[0].get('seed_keyword', ''),
                        'search_volume': keyword_info.get('search_volume', 0),
                        'cpc': keyword_info.get('cpc', 0.0),
                        'competition': keyword_info.get('competition', 0.0)
                    })
            
            # Then look for related keyword items
            for res_item in result:
                if isinstance(res_item, dict) and 'items' in res_item:
                    items = res_item.get('items', [])
                    for item in items:
                        if 'keyword_info' in item:
                            keyword_info = item['keyword_info']
                            related_keywords.append({
                                'keyword': item.get('keyword', ''),
                                'search_volume': keyword_info.get('search_volume', 0),
                                'cpc': keyword_info.get('cpc', 0.0),
                                'competition': keyword_info.get('competition', 0.0)
                            })
            
            # If we found any keywords, return them
            if related_keywords:
                logger.info(f"Successfully extracted {len(related_keywords)} related keywords")
                return related_keywords, True
                
            # If no keywords found, try backup method
            logger.warning(f"No keywords extracted from primary response, trying backup method")
            return fetch_keyword_ideas_backup(keyword, api_login, api_password)
        else:
            error_msg = f"HTTP Error: {response.status_code} - {response.text}"
            logger.error(error_msg)
            return create_default_keywords(keyword), False
    
    except Exception as e:
        error_msg = f"Exception in fetch_related_keywords_dataforseo: {str(e)}"
        logger.error(error_msg)
        logger.error(traceback.format_exc())
        return create_default_keywords(keyword), False

def fetch_keyword_ideas_backup(keyword: str, api_login: str, api_password: str) -> Tuple[List[Dict], bool]:
    """
    Backup method to fetch keyword ideas from DataForSEO Keyword Ideas API
    Returns: keyword_ideas, success_status
    """
    try:
        url = "https://api.dataforseo.com/v3/dataforseo_labs/google/keyword_ideas/live"
        headers = {
            'Content-Type': 'application/json',
        }
        
        # Prepare request data
        post_data = [{
            "keyword": keyword,
            "location_code": 2840,  # USA
            "language_code": "en",
            "include_seed_keyword": True,
            "limit": 20  # Fetch top 20 keyword ideas
        }]
        
        # Log the request
        logger.info(f"Fetching keyword ideas (backup) for: {keyword}")
        
        # Make API request
        response = requests.post(
            url,
            auth=(api_login, api_password),
            headers=headers,
            json=post_data
        )
        
        # Log full response for debugging
        if response.status_code == 200:
            data = response.json()
            logger.info(f"Backup API Response: {json.dumps(data, indent=2)}")
            
            if data.get('status_code') == 20000 and data.get('tasks') and len(data['tasks']) > 0:
                if data['tasks'][0].get('result') and len(data['tasks'][0]['result']) > 0:
                    results = data['tasks'][0]['result']
                    keyword_ideas = []
                    
                    # Process both potential structures
                    # 1. Look for seed keyword data
                    for result in results:
                        if 'seed_keyword_data' in result:
                            seed_data = result['seed_keyword_data']
                            if 'keyword_info' in seed_data:
                                keyword_info = seed_data['keyword_info']
                                keyword_ideas.append({
                                    'keyword': result.get('seed_keyword', ''),
                                    'search_volume': keyword_info.get('search_volume', 0),
                                    'cpc': keyword_info.get('cpc', 0.0),
                                    'competition': keyword_info.get('competition', 0.0)
                                })
                    
                    # 2. Look for items arrays
                    for result in results:
                        if 'items' in result:
                            for item in result['items']:
                                if 'keyword_info' in item:
                                    keyword_info = item['keyword_info']
                                    keyword_ideas.append({
                                        'keyword': item.get('keyword', ''),
                                        'search_volume': keyword_info.get('search_volume', 0),
                                        'cpc': keyword_info.get('cpc', 0.0),
                                        'competition': keyword_info.get('competition', 0.0)
                                    })
                    
                    if keyword_ideas:
                        logger.info(f"Successfully extracted {len(keyword_ideas)} keyword ideas from backup method")
                        return keyword_ideas, True
            
            # Fallback to default keywords
            logger.warning(f"No keyword ideas found in backup method")
            return create_default_keywords(keyword), False
        else:
            error_msg = f"HTTP Error in backup method: {response.status_code} - {response.text}"
            logger.error(error_msg)
            return create_default_keywords(keyword), False
    
    except Exception as e:
        error_msg = f"Exception in fetch_keyword_ideas_backup: {str(e)}"
        logger.error(error_msg)
        logger.error(traceback.format_exc())
        return create_default_keywords(keyword), False

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

def generate_article(keyword: str, semantic_structure: Dict, related_keywords: List[Dict], 
                     serp_features: List[Dict], openai_api_key: str) -> Tuple[str, bool]:
    """
    Generate article using GPT-4o-mini with improved error handling
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
        
        # Generate article
        response = openai.ChatCompletion.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": "You are an expert SEO content writer."},
                {"role": "user", "content": f"""
                Write a comprehensive, SEO-optimized article about "{keyword}".
                
                Use this semantic structure:
                H1: {h1}
                
                Sections:
                {sections_str}
                
                Include these related keywords naturally throughout the text: {related_kw_str}
                
                The top SERP features for this keyword are: {serp_features_str}. Make sure to optimize 
                for these features.
                
                Format the article with proper HTML heading tags (h1, h2, h3) and paragraph tags (p).
                Use bullet points and numbered lists where appropriate.
                
                The article should be comprehensive, factually accurate, and engaging.
                Aim for around 1,500-2,000 words total.
                """}
            ],
            temperature=0.7
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
    Generate internal links for article content using embeddings for similarity
    Returns: article_with_links, links_added, success_status
    """
    try:
        openai.api_key = openai_api_key
        
        # Calculate max links based on word count (10 per 1000 words)
        max_links = min(10, max(1, int(word_count / 1000) * 10))
        
        # Convert pages to simple format for prompt
        pages_str = "\n".join([f"URL: {p['url']}, Title: {p['title']}, Description: {p['description']}" 
                             for p in pages_with_embeddings[:30]])  # Limit to prevent token overflow
        
        # Generate internal links with emphasis on using each URL only once
        response = openai.ChatCompletion.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": "You are an SEO expert specializing in internal linking."},
                {"role": "user", "content": f"""
                Add internal links to this article. Follow these rules:
                
                1. Add no more than {max_links} links total (article has {word_count} words)
                2. Place links approximately every 100 words
                3. Use anchor text of 5-7 words maximum
                4. Only link to relevant pages from the list provided
                5. Format links as HTML <a href="URL">anchor text</a>
                6. Do not modify the article content except to add links
                7. IMPORTANT: USE EACH URL ONLY ONCE - never use the same URL for multiple links
                8. Return the article with added links
                
                Pages available for linking:
                {pages_str}
                
                Article content:
                {article_content}
                
                Also provide a JSON list of links you added in this format:
                [
                    {{
                        "url": "page URL",
                        "anchor_text": "anchor text used",
                        "context": "sentence or phrase containing the link"
                    }}
                ]
                
                Format your response as:
                ARTICLE_WITH_LINKS:
                [article content with links]
                
                LINKS_ADDED:
                [JSON list of links]
                """}
            ],
            temperature=0.3
        )
        
        result = response.choices[0].message.content
        
        # Extract article with links and links added
        article_pattern = re.compile(r'ARTICLE_WITH_LINKS:(.*?)LINKS_ADDED:', re.DOTALL)
        links_pattern = re.compile(r'LINKS_ADDED:(.*)', re.DOTALL)
        
        article_match = article_pattern.search(result)
        links_match = links_pattern.search(result)
        
        article_with_links = article_match.group(1).strip() if article_match else article_content
        
        links_added = []
        if links_match:
            links_json = links_match.group(1).strip()
            # Find JSON array in the text
            json_match = re.search(r'(\[.*\])', links_json, re.DOTALL)
            if json_match:
                try:
                    links_added = json.loads(json_match.group(1))
                except:
                    links_added = []
        
        # Ensure no duplicate URLs (enforce uniqueness)
        if links_added:
            used_urls = set()
            unique_links = []
            
            for link in links_added:
                url = link.get('url', '')
                if url and url not in used_urls:
                    used_urls.add(url)
                    unique_links.append(link)
                else:
                    logger.warning(f"Removing duplicate URL: {url}")
            
            links_added = unique_links
        
        return article_with_links, links_added, True
    
    except Exception as e:
        error_msg = f"Exception in generate_internal_links_with_embeddings: {str(e)}"
        logger.error(error_msg)
        return article_content, [], False

###############################################################################
# 9. Document Generation
###############################################################################

def create_word_document(keyword: str, serp_results: List[Dict], related_keywords: List[Dict],
                        semantic_structure: Dict, article_content: str, meta_title: str, 
                        meta_description: str, internal_links: List[Dict] = None) -> Tuple[BytesIO, bool]:
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
        
        # Section 2: Related Keywords
        doc.add_heading('Related Keywords', level=1)
        
        kw_table = doc.add_table(rows=1, cols=3)
        kw_table.style = 'Table Grid'
        
        # Add header row
        kw_header_cells = kw_table.rows[0].cells
        kw_header_cells[0].text = 'Keyword'
        kw_header_cells[1].text = 'Search Volume'
        kw_header_cells[2].text = 'CPC ($)'
        
        # Add data rows
        for kw in related_keywords:
            row_cells = kw_table.add_row().cells
            row_cells[0].text = kw.get('keyword', '')
            row_cells[1].text = str(kw.get('search_volume', ''))
            row_cells[2].text = f"${kw.get('cpc', 0.0):.2f}"
        
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
                        # Add link as hyperlinked text
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
                    # Fetch SERP results
                    start_time = time.time()
                    organic_results, serp_features, serp_success = fetch_serp_results(
                        keyword, dataforseo_login, dataforseo_password
                    )
                    
                    if serp_success:
                        st.session_state.results['keyword'] = keyword
                        st.session_state.results['organic_results'] = organic_results
                        st.session_state.results['serp_features'] = serp_features
                        
                        # Show SERP results
                        st.subheader("Top 10 Organic Results")
                        df_results = pd.DataFrame(organic_results)
                        st.dataframe(df_results)
                        
                        # Show SERP features
                        st.subheader("SERP Features")
                        df_features = pd.DataFrame(serp_features)
                        st.dataframe(df_features)
                        
                        # Fetch related keywords using DataForSEO
                        st.text("Fetching related keywords...")
                        related_keywords, kw_success = fetch_related_keywords_dataforseo(
                            keyword, dataforseo_login, dataforseo_password
                        )
                        
                        st.session_state.results['related_keywords'] = related_keywords
                        
                        st.subheader("Related Keywords")
                        df_keywords = pd.DataFrame(related_keywords)
                        st.dataframe(df_keywords)
                        
                        if not kw_success:
                            st.info("Using default related keywords. API response was not successful.")
                            
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
                        
                        # Generate article
                        article_content, article_success = generate_article(
                            st.session_state.results['keyword'],
                            st.session_state.results['semantic_structure'],
                            st.session_state.results.get('related_keywords', []),
                            st.session_state.results.get('serp_features', []),
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
                                
                                # Generate internal links
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
                    
                    # Create Word document
                    doc_stream, doc_success = create_word_document(
                        st.session_state.results['keyword'],
                        st.session_state.results['organic_results'],
                        st.session_state.results.get('related_keywords', []),
                        st.session_state.results['semantic_structure'],
                        article_content,
                        meta_title,
                        meta_description,
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
