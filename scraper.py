import requests
from bs4 import BeautifulSoup
import pandas as pd
import json
import re

def scrape_reviews(url):
    """
    Scrapes reviews from klantenvertellen.nl using requests
    Returns a list of dictionaries containing reviewer name, score, and comments
    """
    reviews = []
    
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
    }
    
    try:
        print("Fetching the webpage...")
        response = requests.get(url, headers=headers, timeout=15)
        response.raise_for_status()
        
        soup = BeautifulSoup(response.content, 'html.parser')
        
        # Print page structure for debugging
        print("\n=== Analyzing page structure ===")
        
        # Look for review elements with various possible structures
        review_items = []
        
        # Try finding review containers with common class names
        possible_selectors = [
            soup.find_all('div', class_=re.compile(r'review', re.I)),
            soup.find_all('article', class_=re.compile(r'review', re.I)),
            soup.find_all(class_=re.compile(r'review-item', re.I)),
            soup.find_all(class_=re.compile(r'rating-item', re.I)),
        ]
        
        for items in possible_selectors:
            if items and len(items) > 0:
                review_items = items
                print(f"Found {len(items)} items using selector")
                break
        
        if not review_items:
            print("No review items found with standard selectors")
            print("\nLooking for alternative structures...")
            
            # Try to find any div that might contain review data
            all_divs = soup.find_all('div')
            print(f"Total divs on page: {len(all_divs)}")
            
            # Look for divs with review-related text
            for div in all_divs[:50]:  # Check first 50 divs
                text = div.get_text(strip=True)
                if len(text) > 50 and len(text) < 500:
                    print(f"Potential review div: {div.get('class', 'no-class')}")
                    print(f"Text preview: {text[:100]}...\n")
        
        print(f"\n=== Processing {len(review_items)} reviews ===\n")
        
        for idx, item in enumerate(review_items):
            try:
                # Extract data using multiple possible selectors
                name = extract_text(item, [
                    ('span', 'reviewer-name'),
                    ('span', 'author'),
                    ('div', 'reviewer'),
                    ('p', 'reviewer'),
                    ('div', 'customer-name'),
                ])
                
                score = extract_text(item, [
                    ('span', 'rating'),
                    ('span', 'score'),
                    ('div', 'rating'),
                    ('span', 'stars'),
                    ('div', 'stars'),
                ])
                
                comment = extract_text(item, [
                    ('p', 'review-text'),
                    ('div', 'review-content'),
                    ('div', 'comment'),
                    ('p', 'comment'),
                    ('div', 'review-comment'),
                ])
                
                # If we still haven't found data, get all text from the item
                if not comment:
                    all_text = item.get_text(strip=True)
                    if len(all_text) > 20:
                        comment = all_text
                
                reviews.append({
                    'Reviewer Name': name if name else 'Anonymous',
                    'Score': score if score else 'N/A',
                    'Comments': comment if comment else 'No comment'
                })
                
                print(f"Review {idx + 1}:")
                print(f"  Name: {name if name else 'Anonymous'}")
                print(f"  Score: {score if score else 'N/A'}")
                print(f"  Comment: {(comment[:80] + '...') if comment and len(comment) > 80 else comment if comment else 'No comment'}")
                print()
                
            except Exception as e:
                print(f"Error parsing review {idx}: {e}\n")
                continue
        
        print(f"Successfully extracted {len(reviews)} reviews\n")
        return reviews
    
    except requests.RequestException as e:
        print(f"Error fetching the page: {e}")
        return []
    except Exception as e:
        print(f"Unexpected error: {e}")
        return []

def extract_text(element, selectors):
    """
    Try multiple selectors to extract text from an element
    selectors: list of tuples (tag, class_name)
    """
    for tag, class_name in selectors:
        found = element.find(tag, class_=class_name)
        if found:
            text = found.get_text(strip=True)
            if text:
                return text
    return None

def save_to_excel(reviews, filename='merk_echt_reviews.xlsx'):
    """
    Saves reviews to an Excel file
    """
    if not reviews:
        print("No reviews to save")
        return False
    
    try:
        df = pd.DataFrame(reviews)
        df.to_excel(filename, index=False, sheet_name='Reviews')
        print(f"✓ Reviews saved to '{filename}'")
        print(f"✓ Total reviews exported: {len(df)}")
        print(f"\nFile location: {filename}")
        return True
    except Exception as e:
        print(f"Error saving to Excel: {e}")
        return False

if __name__ == "__main__":
    url = "https://www.klantenvertellen.nl/reviews/1039690/merk_echt"
    
    print("=" * 50)
    print("Merk Echt Review Scraper")
    print("=" * 50)
    print(f"Target URL: {url}\n")
    
    # Scrape reviews
    reviews = scrape_reviews(url)
    
    # Save to Excel
    if reviews:
        save_to_excel(reviews)
    else:
        print("⚠ No reviews were scraped.")
        print("Please manually inspect the website HTML structure")
        print("and update the selectors in the script accordingly.")