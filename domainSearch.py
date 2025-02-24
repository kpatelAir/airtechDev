import whois
import dns.resolver
from googlesearch import search
from fuzzywuzzy import process
from fuzzywuzzy import fuzz
from urllib.parse import urlparse

# --- Enhanced Scoring Based on Multiple Heuristics ---
def score_url(company_name, url):
    domain = urlparse(url).netloc.lower().replace("www.", "")
    score = 0

    # Heuristic 1: Exact domain match
    if company_name.lower() in domain:
        score += 50

    # Heuristic 2: Preferred TLDs
    preferred_tlds = [".com", ".org", ".edu", ".gov", ".net"]
    if any(domain.endswith(tld) for tld in preferred_tlds):
        score += 20

    # Heuristic 3: Fuzzy match for near-exact matches
    similarity_score = fuzz.partial_ratio(company_name.lower(), domain)
    score += similarity_score  # Adds fuzzy match % directly

    # Heuristic 4: Penalize suspicious subdomains (e.g., 'blog.', 'shop.')
    if domain.startswith(("blog.", "shop.", "news.")):
        score -= 10

    return score

# --- Part 1: Google Search to Find Official URL ---
def google_search_for_website(company_name):
    query = f"{company_name} official website"
    try:
        candidates = list(search(query, num_results=10))
        if not candidates:
            return None
        # Score each URL and pick the best
        ranked_candidates = sorted(
            candidates, key=lambda x: score_url(company_name, x), reverse=True
        )
        return ranked_candidates[0]  # Best match based on score
    except Exception as e:
        print(f"Error: {e}")
        return None

# --- Part 2: WHOIS Domain Lookup ---
def verify_domain(domain_url):
    try:
        w = whois.whois(domain_url)
        if w.org or w.domain_name:  # If company/org information exists
            return True
    except Exception:
        return False
    return False

# --- Part 3: DNS Validation ---
def is_valid_domain(domain):
    try:
        dns.resolver.resolve(domain, 'A')  # Look for A records (address)
        return True
    except Exception:
        return False

# --- Part 4: Fuzzy Matching for Near-Matches ---
def fuzzy_match_input(input_site):
    best_match = process.extractOne(input_site)
    return best_match

# --- Main Function ---
def get_official_website(company_name, possible_sites=[]):
    # Step 1: Search Google for official website
    website = google_search_for_website(company_name)
    if website:
        print(f"Initial Google search result: {website}")
        return website
    
    # Step 2: Validate with WHOIS and DNS if domain is accurate
    possible_sites = possible_sites or [f"{company_name.lower()}.com", f"www.{company_name.lower()}.com"]
    for site in possible_sites:
        if verify_domain(site) and is_valid_domain(site):
            print(f"Valid official website found via WHOIS & DNS: {site}")
            return site

# --- Test with a list of companies ---
companies = ["Northrop Grumman - El Segundo", "Boeing", "Lockheed Martin", "Raytheon", "General Dynamics"]

for company in companies:
    print(f"Searching for website of {company}...")
    official_website = get_official_website(company)
    if official_website:
        print(f"Official website for {company}: {official_website}")
    print("\n" + "-"*50 + "\n")
