import requests
from bs4 import BeautifulSoup
import pandas as pd
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options

def scrape_reviews_selenium(url):
    """
    Scrapes reviews from klantenvertellen.nl using Selenium to handle dynamic content
    Returns a list of dictionaries containing reviewer name, score, and comments
    """
    reviews = []
    
    # Set up Chrome options
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    # chrome_options.add_argument("--headless")  # Uncomment for headless mode
    
    try:
        # Initialize the driver
        driver = webdriver.Chrome(options=chrome_options)
        driver.get(url)
        
        # Wait for reviews to load
        WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.CLASS_NAME, "review"))
        )
        
        # Give it some time to fully load
        time.sleep(3)
        
        # Get the page source after JavaScript execution
        soup = BeautifulSoup(driver.page_source, 'html.parser')
        
        # Try multiple possible selectors for review containers
        review_containers = soup.find_all('div', class_='review')
        
        if not review_containers:
            # Try alternative selector
            review_containers = soup.find_all('article', class_='review')
        
        if not review_containers:
            # Try another alternative
            review_containers = soup.find_all(class_='review-item')
        
        print(f"Found {len(review_containers)} review containers")
        
        for idx, review in enumerate(review_containers):
            try:
                # Try multiple selectors for each field
                name = None
                score = None
                comments = None
                
                # Extract reviewer name - try multiple selectors
                name_selectors = [
                    ('span', {'class': 'reviewer-name'}),
                    ('span', {'class': 'author-name'}),
                    ('div', {'class': 'reviewer'}),
                    ('p', {'class': 'reviewer-name'}),
                ]
                
                for tag, attrs in name_selectors:
                    name_element = review.find(tag, attrs)
                    if name_element:
                        name = name_element.text.strip()
                        break
                
                # Extract score - try multiple selectors
                score_selectors = [
                    ('span', {'class': 'rating'}),
                    ('span', {'class': 'score'}),
                    ('div', {'class': 'rating'}),
                    ('span', {'class': 'stars'}),
                ]
                
                for tag, attrs in score_selectors:
                    score_element = review.find(tag, attrs)
                    if score_element:
                        score = score_element.text.strip()
                        break
                
                # Extract comments - try multiple selectors
                comment_selectors = [
                    ('p', {'class': 'review-text'}),
                    ('div', {'class': 'review-content'}),
                    ('p', {'class': 'comment'}),
                    ('div', {'class': 'review-comment'}),
                ]
                
                for tag, attrs in comment_selectors:
                    comment_element = review.find(tag, attrs)
                    if comment_element:
                        comments = comment_element.text.strip()
                        break
                
                # If still not found, print the review HTML for debugging
                if not name or not score or not comments:
                    print(f"\nDebug info for review {idx}:")
                    print(f"HTML snippet: {review.prettify()[:500]}")
                
                reviews.append({
                    'Reviewer Name': name if name else 'Anonymous',
                    'Score': score if score else 'N/A',
                    'Comments': comments if comments else 'No comment'
                })
                
            except Exception as e:
                print(f"Error parsing review {idx}: {e}")
                continue
        
        driver.quit()
        print(f"\nSuccessfully scraped {len(reviews)} reviews")
        return reviews
    
    except Exception as e:
        print(f"Error with Selenium scraper: {e}")
        driver.quit()
        return []

def scrape_reviews_requests(url):
    """
    Fallback scraper using requests library
    """
    reviews = []
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
    }
    
    try:
        response = requests.get(url, headers=headers, timeout=10)
        response.raise_for_status()
        
        soup = BeautifulSoup(response.content, 'html.parser')
        
        # Find all review containers
        review_containers = soup.find_all('div', class_='review')
        
        print(f"Found {len(review_containers)} reviews")
        
        for review in review_containers:
            try:
                name = review.find('span', class_='reviewer-name')
                score = review.find('span', class_='rating')
                comments = review.find('p', class_='review-text')
                
                reviews.append({
                    'Reviewer Name': name.text.strip() if name else 'Anonymous',
                    'Score': score.text.strip() if score else 'N/A',
                    'Comments': comments.text.strip() if comments else ''
                })
            except Exception as e:
                print(f"Error parsing review: {e}")
                continue
        
        return reviews
    
    except requests.RequestException as e:
        print(f"Error fetching the page: {e}")
        return []

def save_to_excel(reviews, filename='merk_echt_reviews.xlsx'):
    """
    Saves reviews to an Excel file
    """
    if not reviews:
        print("No reviews to save")
        return False
    
    df = pd.DataFrame(reviews)
    df.to_excel(filename, index=False, sheet_name='Reviews')
    print(f"✓ Reviews saved to {filename}")
    print(f"Total reviews exported: {len(df)}")
    return True

if __name__ == "__main__":
    url = "https://www.klantenvertellen.nl/reviews/1039690/merk_echt"
    
    print("Starting to scrape reviews...")
    print("Using Selenium for dynamic content loading...\n")
    
    # Try Selenium first (handles dynamic content)
    reviews = scrape_reviews_selenium(url)
    
    # If Selenium fails or returns few results, try requests fallback
    if len(reviews) < 5:
        print("\nTrying fallback scraper...")
        reviews = scrape_reviews_requests(url)
    
    if reviews:
        save_to_excel(reviews)
    else:
        print("No reviews were scraped. Please check the website structure and update selectors accordingly.")