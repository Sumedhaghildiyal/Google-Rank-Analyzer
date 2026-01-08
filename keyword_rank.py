import pandas as pd
import requests
import os
import time
from urllib.parse import urlparse
from dotenv import load_dotenv

# ================== LOAD ENV 
load_dotenv()
API_KEY = os.getenv("SEARCHAPI_KEY")

if not API_KEY:
    raise ValueError("SEARCHAPI_KEY not found. Add it to .env file.")

# ================== GOOGLE SEARCH 
def google_search(keyword, start=0):
    url = "https://www.searchapi.io/api/v1/search"
    params = {
        "engine": "google",
        "q": keyword,
        "location": "Noida, Uttar Pradesh, India",
        "gl": "in",
        "hl": "en",
        "start": start,
        "api_key": API_KEY
    }
    response = requests.get(url, params=params, timeout=30)
    response.raise_for_status()
    return response.json()

# ================== GOOGLE ORGANIC LINKS RANK  

def get_links_rank(keyword, target_url, max_pages=5):
    domain = urlparse(target_url).netloc.replace("www.", "")
    rank_counter = 0

    for page in range(max_pages):
        start = page * 10
        results = google_search(keyword, start=start)
        organic_results = results.get("organic_results", [])

        for result in organic_results:
            rank_counter += 1
            link = result.get("link", "")

            if domain and domain in link:
                return rank_counter

        time.sleep(1)

    return "Not Found"


# ================== GOOGLE PLACES RANK 
#This function tries to find the rank (position) of a 
# business (target URL) in Google Places/local search
#  results for a specific keyword.

def get_places_rank(keyword, target_url, max_pages=5):
    domain = target_url.replace("https://", "").replace("http://", "").replace("www.", "")
    rank_counter = 0 #Keeps track of the position of results in the Google Places results list.

    for page in range(max_pages):
        start = page * 10
        results = google_search(keyword, start=start)
        local_results = results.get("local_results", [])

        if not local_results:
            continue

        for place in local_results:
            rank_counter += 1
            website = place.get("website", "")
            website = website.replace("https://", "").replace("http://", "").replace("www.", "")

            if domain and domain in website:
                return rank_counter

        time.sleep(1)

    return "Not Found"

# ================== MAIN AGENT 
def run_agent():
    # Excel header starts from 2nd row because of "Table 1"
    df = pd.read_excel("keywords.xlsx", header=1)

    # Clean column names
    df.columns = df.columns.str.strip()

    # Validate required columns:Ensures the Excel sheet has both the keyword column and the target page column.
    required_columns = ["Local Keyword Ideas", "Targeted Page"]
    for col in required_columns:
        if col not in df.columns:
            raise ValueError(f"Missing column: {col}")

    places_ranks = [] #will correspond to the Google Places rank for row i
    links_ranks = [] #will store the organic search rank.

    for _, row in df.iterrows():
        keyword = str(row["Local Keyword Ideas"]).strip()
        target_page = str(row["Targeted Page"]).strip()

        #Skip Empty Keywords or URLs
        if not keyword or not target_page:
            places_ranks.append("Skipped")
            links_ranks.append("Skipped")
            continue

        print(f" Checking keyword: {keyword}")

        # Get Places and Links Rank with Error Handling
        try:
            places_rank = get_places_rank(keyword, target_page)
            links_rank = get_links_rank(keyword, target_page)

            places_ranks.append(places_rank)
            links_ranks.append(links_rank)

        except Exception as e:
            print(f"❌ Error for '{keyword}':", e)
            places_ranks.append("Error")
            links_ranks.append("Error")
        #Rate Limiting
        time.sleep(5)

    # Write output
    df["Google Places Rank"] = places_ranks
    df["Google Links Rank"] = links_ranks

    df.to_excel("keyword_ranking_output.xlsx", index=False)
    print("✅ Ranking completed. Output saved as keyword_ranking_output.xlsx")


if __name__ == "__main__":

    run_agent()
