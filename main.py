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
from docx.shared import Pt, Inches
from io import BytesIO
import base64
from typing import List, Dict, Any, Tuple, Optional
import logging

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

def classify_page_type_openai(title: str, snippet: str, openai_api_key: str) -> str:
    """
    Classify the page type using OpenAI
    Returns: page_type (E-commerce, Article, Landing Page, Informational, etc.)
    """
    try:
        openai.api_key = openai_api_key
        
        # Prepare prompt for OpenAI
        prompt = f"""
        Classify the following web page into one of these categories:
        - E-commerce (selling products)
        - Article (blog post, news article, or informational content)
        - Landing Page (focused on conversion/lead generation)
        - Informational (general information, guides, or reference)
        - Tool/App (interactive tool or application)
        - Forum/Community (discussion board or community site)

        Title: {title}
        Snippet: {snippet}

        Return ONLY the category name, nothing else.
        """
        
        response = openai.ChatCompletion.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": "You are a helpful assistant that classifies web pages based on titles and snippets."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.1,
            max_tokens=20
        )
        
        page_type = response.choices[0].message.content.strip()
        
        # Normalize the response to match expected categories
        page_type = page_type.split('\n')[0].strip()
        
        # Handle cases where response doesn't match one of our categories
        valid_types = ["E-commerce", "Article", "Landing Page", "Informational", "Tool/App", "Forum/Community"]
        if page_type not in valid_types:
            # Find closest match
            for valid_type in valid_types:
                if valid_type.lower() in page_type.lower():
                    page_type = valid_type
                    break
            else:
                page_type = "Informational"  # Default if no match
        
        return page_type
    
    except Exception as e:
        logger.error(f"Error classifying page type with OpenAI: {str(e)}")
        return "Informational"  # Default type if there's an error

def fetch_serp_results(keyword: str, api_login: str, api_password: str, openai_api_key: str) -> Tuple[List[Dict], List[Dict], bool]:
    """
    Fetch SERP results from DataForSEO API and classify pages using OpenAI
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
                            # Get title and snippet for OpenAI classification
                            title = item.get('title', '')
                            snippet = item.get('snippet', '')
                            
                            # Classify page type using OpenAI
                            page_type = classify_page_type_openai(title, snippet, openai_api_key)
                            
                            organic_results.append({
                                'url': item.get('url'),
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
        return [], [], False

###############################################################################
# 3. API Integration - DataForSEO for Keywords
###############################################################################

def fetch_related_keywords_dataforseo(keyword: str, api_login: str, api_password: str) -> Tuple[List[Dict], bool]:
    """
    Fetch related keywords from DataForSEO Related Keywords API
    Returns: related_keywords, success_status
    """
    try:
        url = "https://api.dataforseo.com/v3/dataforseo_labs/google/related_keywords/live"
        headers = {
            'Content-Type': 'application/json',
        }
        
        # Prepare request data
        post_data = [{
            "keyword": keyword,
            "location_code": 2840,  # USA
            "language_code": "en",
            "limit": 20  # Fetch top 20 related keywords
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
                
                related_keywords = []
                for item in results.get('items', [])[:20]:  # Limit to top 20
                    related_keywords.append({
                        'keyword': item.get('keyword'),
                        'search_volume': item.get('search_volume', 0),
                        'cpc': item.get('cpc', 0.0),
                        'competition': item.get('competition_index', 0.0)
                    })
                return related_keywords, True
            else:
                error_msg = f"API Error: {data.get('status_message')}"
                logger.error(error_msg)
                return [], False
        else:
            error_msg = f"HTTP Error: {response.status_code} - {response.text}"
            logger.error(error_msg)
            return [], False
    
    except Exception as e:
        error_msg = f"Exception in fetch_related_keywords_dataforseo: {str(e)}"
        logger.error(error_msg)
        return [], False

###############################################################################
# 4. Web Scraping and Content Analysis
###############################################################################

def scrape_webpage(url: str) -> Tuple[str, bool]:
    """
    Scrape webpage content using trafilatura (handles JavaScript better than plain requests)
    Returns: content, success_status
    """
    try:
        # Use trafilatura to download and extract content
        downloaded = trafilatura.fetch_url(url)
        if downloaded:
            content = trafilatura.extract(downloaded, include_comments=False, include_tables=True)
            if content:
                return content, True
            else:
                # Fallback to BeautifulSoup if trafilatura fails
                headers = {
                    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
                }
                response = requests.get(url, headers=headers)
                if response.status_code == 200:
                    soup = BeautifulSoup(response.text, 'html.parser')
                    
                    # Remove script, style elements
                    for script in soup(["script", "style", "header", "footer", "nav"]):
                        script.extract()
                    
                    text = soup.get_text(separator='\n')
                    lines = (line.strip() for line in text.splitlines())
                    chunks = (phrase.strip() for line in lines for phrase in line.split("  "))
                    text = '\n'.join(chunk for chunk in chunks if chunk)
                    
                    return text, True
                else:
                    return "", False
        else:
            return "", False
    
    except Exception as e:
        error_msg = f"Exception in scrape_webpage: {str(e)}"
        logger.error(error_msg)
        return "", False

def extract_headings(url: str) -> Dict[str, List[str]]:
    """
    Extract headings (H1, H2, H3) from a webpage
    """
    try:
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        }
        response = requests.get(url, headers=headers)
        
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
# 5. Embeddings and Semantic Analysis
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
# 6. Content Generation
###############################################################################

def generate_article(keyword: str, semantic_structure: Dict, related_keywords: List[Dict], 
                     serp_features: List[Dict], openai_api_key: str) -> Tuple[str, bool]:
    """
    Generate article using GPT-4o-mini
    Returns: article_content, success_status
    """
    try:
        openai.api_key = openai_api_key
        
        # Prepare section structure
        sections_str = ""
        for section in semantic_structure.get('sections', []):
            sections_str += f"- {section.get('h2')}\n"
            for subsection in section.get('subsections', []):
                sections_str += f"  - {subsection.get('h3')}\n"
        
        # Prepare related keywords
        related_kw_str = ", ".join([kw.get('keyword') for kw in related_keywords[:10]])
        
        # Prepare SERP features
        serp_features_str = ", ".join([f"{feature.get('feature_type')} ({feature.get('count')})" 
                                     for feature in serp_features[:5]])
        
        # Generate article
        response = openai.ChatCompletion.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": "You are an expert SEO content writer."},
                {"role": "user", "content": f"""
                Write a comprehensive, SEO-optimized article about "{keyword}".
                
                Use this semantic structure:
                H1: {semantic_structure.get('h1')}
                
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
# 7. Internal Linking
###############################################################################

def parse_sitemap(sitemap_content: str) -> Tuple[List[Dict], bool]:
    """
    Parse sitemap XML content to extract URLs and metadata
    Returns: pages, success_status
    """
    try:
        soup = BeautifulSoup(sitemap_content, 'xml')
        urls = soup.find_all('url')
        
        pages = []
        for url in urls:
            loc = url.find('loc')
            if loc:
                page_url = loc.text
                # Get title and description by scraping the page
                title, description = get_page_metadata(page_url)
                
                pages.append({
                    'url': page_url,
                    'title': title,
                    'description': description
                })
        
        return pages, True
    
    except Exception as e:
        error_msg = f"Exception in parse_sitemap: {str(e)}"
        logger.error(error_msg)
        return [], False

def get_page_metadata(url: str) -> Tuple[str, str]:
    """
    Get page title and meta description
    """
    try:
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        }
        response = requests.get(url, headers=headers, timeout=5)
        
        if response.status_code == 200:
            soup = BeautifulSoup(response.text, 'html.parser')
            
            title = ""
            title_tag = soup.find('title')
            if title_tag:
                title = title_tag.text.strip()
            
            description = ""
            meta_desc = soup.find('meta', attrs={'name': 'description'})
            if meta_desc and meta_desc.get('content'):
                description = meta_desc['content'].strip()
            
            return title, description
        else:
            return "", ""
    
    except Exception as e:
        logger.error(f"Error fetching metadata for {url}: {str(e)}")
        return "", ""

def generate_internal_links(article_content: str, pages: List[Dict], 
                           openai_api_key: str, word_count: int) -> Tuple[str, List[Dict], bool]:
    """
    Generate internal links for article content
    Returns: article_with_links, links_added, success_status
    """
    try:
        openai.api_key = openai_api_key
        
        # Calculate max links based on word count (10 per 1000 words)
        max_links = min(10, max(1, int(word_count / 1000) * 10))
        
        # Convert pages to simple format for prompt
        pages_str = "\n".join([f"URL: {p['url']}, Title: {p['title']}, Description: {p['description']}" 
                             for p in pages[:30]])  # Limit to prevent token overflow
        
        # Generate internal links
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
                7. Return the article with added links
                
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
        
        return article_with_links, links_added, True
    
    except Exception as e:
        error_msg = f"Exception in generate_internal_links: {str(e)}"
        logger.error(error_msg)
        return article_content, [], False

###############################################################################
# 8. Document Generation
###############################################################################

def create_word_document(keyword: str, serp_results: List[Dict], related_keywords: List[Dict],
                        semantic_structure: Dict, article_content: str, 
                        internal_links: List[Dict] = None) -> Tuple[BytesIO, bool]:
    """
    Create Word document with all components
    Returns: document_stream, success_status
    """
    try:
        doc = Document()
        doc.add_heading(f'SEO Brief: {keyword}', 0)
        
        # Add date
        doc.add_paragraph(f"Generated on: {time.strftime('%Y-%m-%d %H:%M:%S')}")
        
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
        doc.add_heading('Generated Article', level=1)
        
        # Parse HTML content and add to document
        soup = BeautifulSoup(article_content, 'html.parser')
        
        for element in soup.find_all(['h1', 'h2', 'h3', 'p', 'ul', 'ol']):
            if element.name == 'h1':
                doc.add_heading(element.get_text(), level=1)
            elif element.name == 'h2':
                doc.add_heading(element.get_text(), level=2)
            elif element.name == 'h3':
                doc.add_heading(element.get_text(), level=3)
            elif element.name == 'p':
                doc.add_paragraph(element.get_text())
            elif element.name in ['ul', 'ol']:
                for li in element.find_all('li'):
                    doc.add_paragraph(li.get_text(), style='List Bullet' if element.name == 'ul' else 'List Number')
        
        # Section 5: Internal Linking (if provided)
        if internal_links:
            doc.add_heading('Internal Linking Recommendations', level=1)
            
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
        return BytesIO(), False

###############################################################################
# 9. Main Streamlit App
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
                        keyword, dataforseo_login, dataforseo_password, openai_api_key
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
                        
                        if kw_success:
                            st.session_state.results['related_keywords'] = related_keywords
                            
                            st.subheader("Related Keywords")
                            df_keywords = pd.DataFrame(related_keywords)
                            st.dataframe(df_keywords)
                            
                            st.success(f"SERP analysis completed in {format_time(time.time() - start_time)}")
                        else:
                            st.error("Failed to fetch related keywords. Please check your DataForSEO API credentials.")
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
                            
                            if success:
                                scraped_contents.append({
                                    'url': result['url'],
                                    'title': result['title'],
                                    'content': content
                                })
                                
                                # Also extract headings
                                headings = extract_headings(result['url'])
                                scraped_contents[-1]['headings'] = headings
                            
                            progress_bar.progress((i + 1) / len(st.session_state.results['organic_results']))
                        
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
            if st.button("Generate Article"):
                if not openai_api_key:
                    st.error("Please enter OpenAI API key")
                else:
                    with st.spinner("Generating article..."):
                        start_time = time.time()
                        
                        article_content, article_success = generate_article(
                            st.session_state.results['keyword'],
                            st.session_state.results['semantic_structure'],
                            st.session_state.results.get('related_keywords', []),
                            st.session_state.results.get('serp_features', []),
                            openai_api_key
                        )
                        
                        if article_success:
                            st.session_state.results['article_content'] = article_content
                            
                            st.subheader("Generated Article")
                            st.markdown(article_content, unsafe_allow_html=True)
                            
                            st.success(f"Article generation completed in {format_time(time.time() - start_time)}")
                        else:
                            st.error("Failed to generate article")
            
            # Show previously generated article if available
            if 'article_content' in st.session_state.results:
                st.subheader("Previously Generated Article")
                st.markdown(st.session_state.results['article_content'], unsafe_allow_html=True)
    
    # Tab 4: Internal Linking
    with tabs[3]:
        st.header("Internal Linking")
        
        if 'article_content' not in st.session_state.results:
            st.warning("Please generate an article first (in the 'Article Generation' tab)")
        else:
            st.write("Upload a sitemap XML file to generate internal link suggestions:")
            sitemap_file = st.file_uploader("Upload Sitemap XML", type=['xml'])
            
            sitemap_url = st.text_input("Or enter sitemap URL (e.g., https://example.com/sitemap.xml)")
            
            if st.button("Process Sitemap"):
                if not openai_api_key:
                    st.error("Please enter OpenAI API key")
                elif not sitemap_file and not sitemap_url:
                    st.error("Please either upload a sitemap file or enter a sitemap URL")
                else:
                    with st.spinner("Processing sitemap and generating internal links..."):
                        start_time = time.time()
                        
                        # Process sitemap
                        if sitemap_file:
                            sitemap_content = sitemap_file.read().decode('utf-8')
                        else:
                            try:
                                response = requests.get(sitemap_url)
                                if response.status_code == 200:
                                    sitemap_content = response.text
                                else:
                                    st.error(f"Failed to fetch sitemap: HTTP {response.status_code}")
                                    sitemap_content = None
                            except Exception as e:
                                st.error(f"Error fetching sitemap: {str(e)}")
                                sitemap_content = None
                        
                        if sitemap_content:
                            pages, sitemap_success = parse_sitemap(sitemap_content)
                            
                            if sitemap_success and pages:
                                st.session_state.results['sitemap_pages'] = pages
                                
                                # Count words in the article
                                article_content = st.session_state.results['article_content']
                                word_count = len(re.findall(r'\w+', article_content))
                                
                                # Generate internal links
                                article_with_links, links_added, links_success = generate_internal_links(
                                    article_content, pages, openai_api_key, word_count
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
                                st.error("Failed to parse sitemap or no pages found")
            
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
                    
                    # Create Word document
                    doc_stream, doc_success = create_word_document(
                        st.session_state.results['keyword'],
                        st.session_state.results['organic_results'],
                        st.session_state.results.get('related_keywords', []),
                        st.session_state.results['semantic_structure'],
                        article_content,
                        internal_links
                    )
                    
                    if doc_success:
                        st.session_state.results['doc_stream'] = doc_stream
                        
                        st.download_button(
                            label="Download SEO Brief",
                            data=doc_stream,
                            file_name=f"seo_brief_{st.session_state.results['keyword'].replace(' ', '_')}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
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
                ("Generated Article", 'article_content' in st.session_state.results),
                ("Internal Linking", 'article_with_links' in st.session_state.results)
            ]
            
            for component, status in components:
                st.write(f"**{component}:** {'✅ Completed' if status else '❌ Not Completed'}")
            
            # Display download button if available
            if 'doc_stream' in st.session_state.results:
                st.subheader("Download SEO Brief")
                st.download_button(
                    label="Download SEO Brief",
                    data=st.session_state.results['doc_stream'],
                    file_name=f"seo_brief_{st.session_state.results['keyword'].replace(' ', '_')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

if __name__ == "__main__":
    main()
