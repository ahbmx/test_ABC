Understood! You need to use **username and password** (basic authentication) instead of API tokens. Here's the updated comprehensive cheat sheet:

---

## 🚀 Setup with Username/Password Authentication

### Installation

```bash
pip install atlassian-python-api pandas requests
```

### Hard-Coded Username/Password Authentication

```python
from atlassian import Confluence
import pandas as pd
import requests
from requests.auth import HTTPBasicAuth
import base64
import json

# ============================================================
# HARD-CODED USERNAME & PASSWORD (As Required)
# ⚠️ WARNING: Never commit this to public repositories!
# ============================================================

confluence_url = "https://your-domain.atlassian.net/wiki"
username = "your.email@example.com"
password = "your_actual_password"  # Your Confluence login password

# Initialize Confluence client with username/password
confluence = Confluence(
    url=confluence_url,
    username=username,
    password=password  # Using actual password instead of API token
)

# For direct REST API calls (useful for advanced operations)
auth = HTTPBasicAuth(username, password)
headers = {
    "Content-Type": "application/json",
    "X-Atlassian-Token": "no-check"  # Required for file uploads
}
```

> ⚠️ **Important**: 
> - If using **Confluence Cloud** with username/password, you may need to use an [App Password](https://support.atlassian.com/atlassian-account/docs/manage-api-tokens-for-your-atlassian-account/) if your organization has MFA enabled
> - For **Confluence Server/Data Center**, standard username/password works directly
> - Never hard-code credentials in production code

---

## 📄 Creating & Updating Pages

### Page Styling with Storage Format

Confluence uses **Storage Format** (XML-based) for rich content. Here are styling patterns:

```python
def create_styled_page(space_key, title, parent_id=None):
    """Create a page with custom styling"""
    
    # Storage format with CSS styling
    styled_content = """
    <ac:structured-macro ac:name="info">
        <ac:rich-text-body>
            <p><strong>📌 Important:</strong> This page was automated!</p>
        </ac:rich-text-body>
    </ac:structured-macro>
    
    <h1 style="color: #0052CC; border-bottom: 2px solid #0052CC;">
        Main Heading
    </h1>
    
    <h2>Subheading with <span style="color: #FF5630;">colored text</span></h2>
    
    <p>
        <span style="background-color: #EBECF0; padding: 10px; border-radius: 5px;">
            Highlighted text with background
        </span>
    </p>
    
    <ac:structured-macro ac:name="panel">
        <ac:parameter ac:name="title">Info Panel</ac:parameter>
        <ac:parameter ac:name="borderColor">#00B8D9</ac:parameter>
        <ac:parameter ac:name="bgColor">#E6F7FF</ac:parameter>
        <ac:rich-text-body>
            <p>Content inside a styled panel</p>
        </ac:rich-text-body>
    </ac:structured-macro>
    
    <ac:structured-macro ac:name="code">
        <ac:parameter ac:name="language">python</ac:parameter>
        <ac:plain-text-body><![CDATA[
            print("Code block with syntax highlighting")
        ]]></ac:plain-text-body>
    </ac:structured-macro>
    """
    
    page = confluence.create_page(
        space=space_key,
        title=title,
        body=styled_content,
        parent_id=parent_id,
        type='page',
        representation='storage'
    )
    return page['id']

def update_page_with_styling(page_id, new_title, new_content):
    """Update existing page with styled content"""
    page = confluence.get_page_by_id(page_id, expand='version')
    current_version = page['version']['number']
    
    result = confluence.update_page(
        page_id=page_id,
        title=new_title,
        body=new_content,
        version=current_version + 1,
        representation='storage'
    )
    print(f"✅ Page updated: {confluence_url}/pages/{page_id}")
    return result
```

---

## 📊 Creating Tables from DataFrames

### Basic DataFrame to Table

```python
def dataframe_to_confluence_table(df, include_header=True):
    """
    Convert pandas DataFrame to Confluence storage format table.
    
    Args:
        df: pandas DataFrame
        include_header: Whether to include column headers
    """
    rows = []
    
    # Add header row
    if include_header:
        header_cells = ''.join([f'<th><p><strong>{col}</strong></p></th>' for col in df.columns])
        rows.append(f'<tr>{header_cells}</tr>')
    
    # Add data rows with styling
    for idx, row in df.iterrows():
        cells = []
        for col in df.columns:
            value = str(row[col])
            
            # Apply conditional styling
            if 'error' in value.lower() or 'fail' in value.lower():
                cell_style = 'style="background-color: #FFEBE6; color: #DE350B;"'
            elif 'success' in value.lower() or 'pass' in value.lower():
                cell_style = 'style="background-color: #E3FCEF; color: #006644;"'
            else:
                cell_style = ''
            
            cells.append(f'<td {cell_style}><p>{value}</p></td>')
        
        # Alternate row background color
        row_style = 'style="background-color: #F4F5F7;"' if idx % 2 == 0 else ''
        rows.append(f'<tr {row_style}>{"".join(cells)}</tr>')
    
    # Wrap in table with borders
    table_html = f'''
    <table style="border-collapse: collapse; width: 100%;">
        {"".join(rows)}
    </table>
    '''
    
    return table_html

# Example usage
df = pd.DataFrame({
    'Test Case': ['TC-001', 'TC-002', 'TC-003'],
    'Status': ['Passed', 'Failed', 'Passed'],
    'Duration (s)': [1.2, 3.4, 2.1]
})

table_content = dataframe_to_confluence_table(df, include_header=True)

# Create or update page with table
page_id = confluence.create_page(
    space='DEV',
    title='Test Results - DataFrame Report',
    body=f'<h1>Test Execution Report</h1>{table_content}'
)
```

### Advanced Table with Merged Cells & Specific Dimensions

```python
def create_advanced_table_with_styling():
    """Create a table with specific cell dimensions and merged cells"""
    
    table_html = '''
    <table style="border-collapse: collapse; width: 100%; table-layout: fixed;">
        <colgroup>
            <col style="width: 20%;">
            <col style="width: 50%;">
            <col style="width: 30%;">
        </colgroup>
        <thead>
            <tr style="background-color: #172B4D; color: white;">
                <th style="padding: 12px; text-align: left;">Category</th>
                <th style="padding: 12px; text-align: left;">Description</th>
                <th style="padding: 12px; text-align: center;">Status</th>
            </tr>
        </thead>
        <tbody>
            <tr>
                <td rowspan="2" style="vertical-align: middle; background-color: #F4F5F7;">
                    Merged Cell
                </td>
                <td style="padding: 10px;">Row 1 content here</td>
                <td style="text-align: center;">✅</td>
            </tr>
            <tr>
                <td style="padding: 10px;">Row 2 content here</td>
                <td style="text-align: center;">⚠️</td>
            </tr>
        </tbody>
    </table>
    '''
    return table_html
```

---

## 📎 Uploading Files & Embedding Images

### Uploading File Attachments

```python
def upload_file_to_page(page_id, file_path, comment=None):
    """
    Upload a file attachment to a Confluence page.
    
    Args:
        page_id: The Confluence page ID
        file_path: Path to the file to upload
        comment: Optional attachment comment
    """
    with open(file_path, 'rb') as file:
        files = {'file': (os.path.basename(file_path), file, 'application/octet-stream')}
        
        url = f"{confluence_url}/rest/api/content/{page_id}/child/attachment"
        
        response = requests.post(
            url,
            auth=auth,
            headers=headers,
            files=files,
            data={'comment': comment} if comment else None
        )
        
        if response.status_code == 200:
            print(f"✅ File uploaded: {os.path.basename(file_path)}")
            return response.json()
        else:
            print(f"❌ Upload failed: {response.text}")
            return None

# Example
upload_file_to_page(page_id=12345, file_path="report.pdf", comment="Q1 Report")
```

### Embedding Images with Specific Dimensions

```python
def upload_and_embed_image(page_id, image_path, width=None, height=None, alignment='center'):
    """
    Upload image and get storage format for embedding with specific dimensions.
    
    Args:
        page_id: Confluence page ID
        image_path: Path to image file
        width: Desired width in pixels (e.g., 400)
        height: Desired height in pixels (e.g., 300)
        alignment: 'left', 'center', or 'right'
    """
    # Upload the image
    with open(image_path, 'rb') as img_file:
        files = {'file': (os.path.basename(image_path), img_file, 'image/png')}
        
        upload_url = f"{confluence_url}/rest/api/content/{page_id}/child/attachment"
        response = requests.post(upload_url, auth=auth, headers=headers, files=files)
        
        if response.status_code != 200:
            raise Exception(f"Image upload failed: {response.text}")
        
        attachment_id = response.json()['results'][0]['id']
    
    # Get attachment details to get download URL
    attachment_url = f"{confluence_url}/rest/api/content/{page_id}/child/attachment/{attachment_id}"
    attachment_info = requests.get(attachment_url, auth=auth).json()
    download_link = f"{confluence_url}{attachment_info['_links']['download']}"
    
    # Create storage format with specific image dimensions
    align_class = {
        'left': 'align-left',
        'center': 'align-center',
        'right': 'align-right'
    }.get(alignment, 'align-center')
    
    width_attr = f'width="{width}"' if width else ''
    height_attr = f'height="{height}"' if height else ''
    
    # Confluence storage format for embedded image
    image_macro = f'''
    <ac:image {width_attr} {height_attr}>
        <ri:attachment ri:filename="{os.path.basename(image_path)}" />
    </ac:image>
    '''
    
    return image_macro

def embed_image_in_table_cell(page_id, image_path, row, col, width=None, height=None):
    """
    Embed an image into a specific table cell.
    """
    # First, create a table with placeholder
    table_html = f'''
    <table>
        <tr>
            <td>Cell 1</td>
            <td id="image-cell-{row}-{col}">
                <!-- IMAGE WILL BE EMBEDDED HERE -->
            </td>
        </tr>
    </table>
    '''
    
    # Upload image and get macro
    image_macro = upload_and_embed_image(page_id, image_path, width, height)
    
    # Replace placeholder with actual image
    final_html = table_html.replace(f'<!-- IMAGE WILL BE EMBEDDED HERE -->', image_macro)
    
    return final_html
```

### Complete Example: Image in Specific Table Cell

```python
def create_page_with_images_in_table(space_key, title):
    """
    Create a page with a table containing images in specific cells.
    """
    # Upload images first (they need a page ID)
    temp_page = confluence.create_page(
        space=space_key,
        title="TEMP_For_Images",
        body="<p>Temporary page for images</p>"
    )
    temp_page_id = temp_page['id']
    
    # Upload images and get their macros
    image1_macro = upload_and_embed_image(temp_page_id, "chart1.png", width=300, height=200)
    image2_macro = upload_and_embed_image(temp_page_id, "chart2.png", width=400, height=300)
    
    # Create table with images in specific cells
    table_with_images = f'''
    <h2>Dashboard Report</h2>
    <table style="border-collapse: collapse; width: 100%;">
        <thead>
            <tr style="background-color: #42526E; color: white;">
                <th style="padding: 10px; width: 30%;">Metric</th>
                <th style="padding: 10px; width: 70%;">Chart</th>
            </tr>
        </thead>
        <tbody>
            <tr>
                <td style="padding: 10px; vertical-align: middle;">
                    <strong>Revenue Q1 2024</strong>
                    <p>💰 $1.2M</p>
                </td>
                <td style="text-align: center; padding: 20px;">
                    {image1_macro}
                </td>
            </tr>
            <tr>
                <td style="padding: 10px; vertical-align: middle;">
                    <strong>User Growth</strong>
                    <p>📈 +45% YoY</p>
                </td>
                <td style="text-align: center; padding: 20px;">
                    {image2_macro}
                </td>
            </tr>
        </tbody>
    </table>
    '''
    
    # Update the real page with the table
    final_page = confluence.update_page(
        page_id=temp_page_id,
        title=title,
        body=table_with_images,
        version=2,
        representation='storage'
    )
    
    print(f"✅ Page created with images in table cells!")
    return final_page

# Usage
create_page_with_images_in_table("DEV", "Dashboard - Q1 2024")
```

---

## 🎨 Additional Styling Macros

```python
def create_info_box(message, type='info'):
    """Create styled info boxes"""
    styles = {
        'info': {'color': '#0052CC', 'bg': '#DEEBFF', 'icon': 'ℹ️'},
        'success': {'color': '#006644', 'bg': '#E3FCEF', 'icon': '✅'},
        'warning': {'color': '#FF8B00', 'bg': '#FFFAE6', 'icon': '⚠️'},
        'error': {'color': '#DE350B', 'bg': '#FFEBE6', 'icon': '❌'}
    }
    
    style = styles.get(type, styles['info'])
    
    return f'''
    <ac:structured-macro ac:name="info">
        <ac:parameter ac:name="title">{style['icon']} {type.upper()}</ac:parameter>
        <ac:rich-text-body>
            <p style="color: {style['color']}; background: {style['bg']}; padding: 15px;">
                {message}
            </p>
        </ac:rich-text-body>
    </ac:structured-macro>
    '''
```

---

## 🔍 Utility Functions

```python
def get_page_id_by_title(space_key, title):
    """Find page ID by title"""
    pages = confluence.get_all_pages_from_space(space_key, start=0, limit=50)
    for page in pages:
        if page['title'] == title:
            return page['id']
    return None

def delete_all_attachments(page_id):
    """Remove all attachments from a page"""
    url = f"{confluence_url}/rest/api/content/{page_id}/child/attachment"
    attachments = requests.get(url, auth=auth).json()
    
    for att in attachments.get('results', []):
        delete_url = f"{url}/{att['id']}"
        requests.delete(delete_url, auth=auth)
        print(f"Deleted: {att['title']}")
```

---

## ⚠️ Important Notes for Username/Password Authentication

1. **Confluence Cloud**: If using cloud version, you may need to:
   - Enable "Less secure apps" (not recommended)
   - Use an [App Password](https://support.atlassian.com/atlassian-account/docs/manage-app-passwords-for-your-atlassian-account/) instead of your main password
   - Configure [CORS settings](https://developer.atlassian.com/cloud/confluence/rest-api-intro/#cors)

2. **Confluence Server/Data Center**: Username/password works directly

3. **Security Best Practices** (even though hard-coding is required):
   ```python
   # At minimum, obfuscate slightly or use environment variables
   import os
   username = os.getenv('CONFLUENCE_USER', 'hardcoded_user@example.com')
   password = os.getenv('CONFLUENCE_PASS', 'hardcoded_password')
   ```

4. **Troubleshooting Authentication**:
   ```python
   # Test authentication
   try:
       spaces = confluence.get_all_spaces()
       print("✅ Authentication successful!")
   except Exception as e:
       print(f"❌ Auth failed: {e}")
   ```

This cheat sheet covers all your requirements with hard-coded username/password authentication! 🚀





Here's a comprehensive function to list all href links in a Confluence page, match them against a string pattern, and extract the page IDs and parent information:

```python
from atlassian import Confluence
import re
from urllib.parse import urlparse, parse_qs
from typing import List, Dict, Tuple, Optional

# Initialize connection (using your hard-coded credentials)
confluence = Confluence(
    url=confluence_url,
    username=username,
    password=password
)

def extract_page_id_from_url(url: str) -> Optional[int]:
    """
    Extract page ID from various Confluence URL formats.
    
    Handles formats:
    - https://domain/wiki/spaces/SPACEKEY/pages/123456789/Page+Title
    - https://domain/wiki/pages/123456789/Page+Title
    - https://domain/wiki/pages/viewpage.action?pageId=123456789
    - /pages/123456789/Page+Title (relative URLs)
    """
    if not url:
        return None
    
    # Pattern 1: /pages/123456789/ (absolute or relative)
    pattern1 = r'/pages/(\d+)(?:/|$)'
    match = re.search(pattern1, url)
    if match:
        return int(match.group(1))
    
    # Pattern 2: pageId=123456789 in query parameters
    pattern2 = r'[?&]pageId=(\d+)'
    match = re.search(pattern2, url)
    if match:
        return int(match.group(1))
    
    # Pattern 3: /spaces/KEY/pages/123456789
    pattern3 = r'/spaces/[^/]+/pages/(\d+)'
    match = re.search(pattern3, url)
    if match:
        return int(match.group(1))
    
    return None

def get_page_parent_info(page_id: int) -> Dict:
    """
    Get parent information for a given page ID.
    
    Returns dictionary with parent_id, parent_title, and parent_url.
    """
    try:
        page = confluence.get_page_by_id(page_id, expand='space,version,ancestors')
        
        # Get ancestors (parent chain)
        ancestors = page.get('ancestors', [])
        
        if ancestors:
            # The last ancestor is the immediate parent
            parent = ancestors[-1]
            parent_info = {
                'parent_id': parent['id'],
                'parent_title': parent['title'],
                'parent_url': f"{confluence_url}{parent.get('_links', {}).get('webui', '')}"
            }
        else:
            parent_info = {
                'parent_id': None,
                'parent_title': 'No parent (top-level page)',
                'parent_url': None
            }
        
        return parent_info
    
    except Exception as e:
        print(f"❌ Error getting parent for page {page_id}: {e}")
        return {
            'parent_id': None,
            'parent_title': f'Error: {str(e)}',
            'parent_url': None
        }

def find_and_match_hrefs(
    page_id: int,
    match_string: str,
    include_external_links: bool = False
) -> List[Dict]:
    """
    Extract all href links from a Confluence page, match them against a string,
    and retrieve page IDs and parent information.
    
    Args:
        page_id: The Confluence page ID to scan
        match_string: String pattern to match in href URLs
        include_external_links: Whether to include external (non-Confluence) links
    
    Returns:
        List of dictionaries containing matching link information
    """
    
    # Get the page content in storage format
    try:
        page = confluence.get_page_by_id(
            page_id, 
            expand='body.storage,version,space,ancestors'
        )
        content = page['body']['storage']['value']
        page_title = page['title']
        page_url = f"{confluence_url}/pages/{page_id}"
        
        print(f"🔍 Scanning page: '{page_title}' (ID: {page_id})")
        print(f"📄 Page URL: {page_url}")
        print(f"🎯 Looking for links matching: '{match_string}'")
        print("-" * 80)
        
    except Exception as e:
        print(f"❌ Failed to retrieve page {page_id}: {e}")
        return []
    
    # Find all href links using regex
    # Pattern matches href="..." or href='...'
    href_pattern = r'href=["\']([^"\']+)["\']'
    all_hrefs = re.findall(href_pattern, content)
    
    print(f"📊 Found {len(all_hrefs)} total links on the page")
    
    # Filter and process matching links
    matched_links = []
    
    for href in all_hrefs:
        # Skip empty links or anchors
        if not href or href == '#' or href.startswith('javascript:'):
            continue
        
        # Check if link matches the string pattern
        if match_string.lower() in href.lower():
            print(f"\n✅ Found matching link: {href}")
            
            # Determine if it's a Confluence internal link
            is_confluence_link = (
                'confluence' in href.lower() or 
                '/pages/' in href or 
                '/spaces/' in href or
                'pageId=' in href
            )
            
            link_info = {
                'original_href': href,
                'is_confluence_link': is_confluence_link,
                'page_title': None,
                'page_id': None,
                'parent_info': None,
                'space_key': None
            }
            
            # Extract page ID from Confluence links
            if is_confluence_link or include_external_links:
                linked_page_id = extract_page_id_from_url(href)
                
                if linked_page_id:
                    link_info['page_id'] = linked_page_id
                    
                    # Try to get page details
                    try:
                        linked_page = confluence.get_page_by_id(
                            linked_page_id, 
                            expand='version,space,ancestors'
                        )
                        link_info['page_title'] = linked_page.get('title', 'Unknown')
                        link_info['space_key'] = linked_page.get('space', {}).get('key', 'Unknown')
                        
                        # Get parent information for the linked page
                        parent_info = get_page_parent_info(linked_page_id)
                        link_info['parent_info'] = parent_info
                        
                        print(f"   📄 Page Title: {link_info['page_title']}")
                        print(f"   🆔 Page ID: {linked_page_id}")
                        print(f"   📂 Space: {link_info['space_key']}")
                        if parent_info['parent_id']:
                            print(f"   👪 Parent: '{parent_info['parent_title']}' (ID: {parent_info['parent_id']})")
                        else:
                            print(f"   👪 Parent: {parent_info['parent_title']}")
                            
                    except Exception as e:
                        print(f"   ⚠️ Could not fetch page details: {e}")
                        link_info['page_title'] = 'Unable to fetch (maybe deleted or no access)'
                else:
                    # External or non-Confluence link
                    if include_external_links:
                        link_info['page_title'] = 'External Link'
                        link_info['parent_info'] = {'parent_title': 'N/A', 'parent_id': None}
                        print(f"   🔗 External link (no page ID)")
            
            matched_links.append(link_info)
    
    # Summary
    print("\n" + "=" * 80)
    print(f"📈 SUMMARY: Found {len(matched_links)} links matching '{match_string}'")
    print("=" * 80)
    
    return matched_links

# ============================================================
# ENHANCED VERSIONS WITH ADDITIONAL FEATURES
# ============================================================

def find_hrefs_with_regex_pattern(
    page_id: int,
    regex_pattern: str,
    case_sensitive: bool = False
) -> List[Dict]:
    """
    Advanced version: Use regex pattern for matching hrefs.
    
    Args:
        page_id: Confluence page ID
        regex_pattern: Regular expression pattern to match
        case_sensitive: Whether to match case sensitively
    """
    flags = 0 if case_sensitive else re.IGNORECASE
    compiled_pattern = re.compile(regex_pattern, flags)
    
    page = confluence.get_page_by_id(page_id, expand='body.storage')
    content = page['body']['storage']['value']
    
    href_pattern = r'href=["\']([^"\']+)["\']'
    all_hrefs = re.findall(href_pattern, content)
    
    matched_links = []
    
    for href in all_hrefs:
        if compiled_pattern.search(href):
            linked_page_id = extract_page_id_from_url(href)
            link_data = {
                'href': href,
                'page_id': linked_page_id
            }
            
            if linked_page_id:
                try:
                    linked_page = confluence.get_page_by_id(linked_page_id)
                    link_data['title'] = linked_page.get('title')
                    link_data['parent_info'] = get_page_parent_info(linked_page_id)
                except:
                    link_data['title'] = 'Unknown'
            
            matched_links.append(link_data)
    
    return matched_links

def get_all_links_with_parents_recursive(
    start_page_id: int,
    match_string: str,
    max_depth: int = 3,
    visited_pages: set = None
) -> List[Dict]:
    """
    Recursively scan pages and their linked pages for matching hrefs.
    
    Args:
        start_page_id: Starting page ID
        match_string: String to match in hrefs
        max_depth: Maximum recursion depth
        visited_pages: Set of already visited pages (to avoid cycles)
    """
    if visited_pages is None:
        visited_pages = set()
    
    if max_depth <= 0 or start_page_id in visited_pages:
        return []
    
    visited_pages.add(start_page_id)
    
    results = []
    
    # Find matches on current page
    current_matches = find_and_match_hrefs(start_page_id, match_string)
    
    for match in current_matches:
        match['source_page_id'] = start_page_id
        match['depth_remaining'] = max_depth - 1
        results.append(match)
        
        # Recursively scan linked pages if they have page IDs
        if match.get('page_id') and max_depth > 1:
            child_results = get_all_links_with_parents_recursive(
                match['page_id'],
                match_string,
                max_depth - 1,
                visited_pages
            )
            results.extend(child_results)
    
    return results

def export_matches_to_dataframe(matched_links: List[Dict]) -> pd.DataFrame:
    """
    Export matched links to a pandas DataFrame for easy analysis.
    """
    rows = []
    
    for link in matched_links:
        row = {
            'Original_Href': link['original_href'],
            'Is_Confluence_Link': link['is_confluence_link'],
            'Page_Title': link.get('page_title'),
            'Page_ID': link.get('page_id'),
            'Space_Key': link.get('space_key'),
            'Parent_ID': link.get('parent_info', {}).get('parent_id'),
            'Parent_Title': link.get('parent_info', {}).get('parent_title'),
            'Parent_URL': link.get('parent_info', {}).get('parent_url')
        }
        rows.append(row)
    
    return pd.DataFrame(rows)

# ============================================================
# USAGE EXAMPLES
# ============================================================

def example_usage():
    """Demonstrate how to use the functions"""
    
    # Example 1: Basic usage - find all links containing "documentation"
    page_id_to_scan = 123456789  # Replace with your page ID
    matches = find_and_match_hrefs(
        page_id=page_id_to_scan,
        match_string="documentation",
        include_external_links=False
    )
    
    # Print results
    for match in matches:
        print(f"\n🔗 Link: {match['original_href']}")
        if match['page_id']:
            print(f"   → Page ID: {match['page_id']}")
            print(f"   → Title: {match['page_title']}")
            print(f"   → Parent: {match['parent_info']['parent_title']}")
    
    # Example 2: Export to DataFrame
    df = export_matches_to_dataframe(matches)
    print("\n📊 DataFrame Summary:")
    print(df[['Page_Title', 'Page_ID', 'Parent_Title']].head())
    
    # Example 3: Recursive scanning
    all_matches = get_all_links_with_parents_recursive(
        start_page_id=page_id_to_scan,
        match_string="api",
        max_depth=2
    )
    print(f"\n🔄 Found {len(all_matches)} matches across linked pages")
    
    # Example 4: Using regex pattern
    import re
    regex_matches = find_hrefs_with_regex_pattern(
        page_id=page_id_to_scan,
        regex_pattern=r'/pages/\d+',  # Find all internal page links
        case_sensitive=False
    )
    print(f"\n🎯 Found {len(regex_matches)} internal page links using regex")

# Run the example
if __name__ == "__main__":
    # Configure your credentials first
    confluence_url = "https://your-domain.atlassian.net/wiki"
    username = "your.email@example.com"
    password = "your_password"
    
    # Initialize connection
    confluence = Confluence(url=confluence_url, username=username, password=password)
    
    # Example: Scan a specific page
    target_page_id = 123456789  # Replace with actual page ID
    matching_links = find_and_match_hrefs(
        page_id=target_page_id,
        match_string="https://",  # Find all HTTP/HTTPS links
        include_external_links=True
    )
    
    # Display results in a table format
    print("\n📋 Final Results Table:")
    print("-" * 100)
    for idx, link in enumerate(matching_links, 1):
        print(f"{idx}. HREF: {link['original_href'][:80]}...")
        if link['page_id']:
            print(f"   → Page ID: {link['page_id']} | Title: {link['page_title']}")
            print(f"   → Parent: {link['parent_info']['parent_title']}")
        print()
```

## Key Features of the Function:

1. **Extracts all href links** from any Confluence page content
2. **Matches links** against a specified string pattern (case-insensitive)
3. **Extracts page IDs** from multiple Confluence URL formats:
   - `/pages/123456789/Page+Title`
   - `/spaces/KEY/pages/123456789`
   - `?pageId=123456789` query parameters
   - Relative and absolute URLs
4. **Retrieves parent information** for each linked page including:
   - Parent page ID
   - Parent page title
   - Parent page URL
5. **Handles errors gracefully** when pages are inaccessible or deleted
6. **Provides detailed output** with progress indicators

## Additional Utility Functions Included:

- **`find_hrefs_with_regex_pattern`**: For complex pattern matching using regex
- **`get_all_links_with_parents_recursive`**: Recursively scan linked pages
- **`export_matches_to_dataframe`**: Convert results to pandas DataFrame for analysis

## Sample Output:
```
🔍 Scanning page: 'Documentation Hub' (ID: 123456789)
📄 Page URL: https://your-domain/wiki/pages/123456789
🎯 Looking for links matching: 'api'
📊 Found 45 total links on the page

✅ Found matching link: /pages/987654321/REST+API+Guide
   📄 Page Title: REST API Guide
   🆔 Page ID: 987654321
   📂 Space: DEV
   👪 Parent: 'API Documentation' (ID: 987654320)
```

This function gives you complete visibility into all links on a Confluence page, their targets, and their parent relationships!
