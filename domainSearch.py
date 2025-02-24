import whois
import dns.resolver
from googlesearch import search
from fuzzywuzzy import process

# --- Part 1: Google Search to Find Official URL ---
def google_search_for_website(company_name):
    query = f"{company_name} official website"
    # Retrieve top 5 results
    for result in search(query, num_results=5):
        if company_name.lower() in result.lower():  # prioritize results with company name
            return result
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
def fuzzy_match_input(input_site, possible_sites):
    best_match = process.extractOne(input_site, possible_sites)
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

    # Step 3: Fuzzy matching for near-matches (if provided a list)
    if possible_sites:
        best_match = fuzzy_match_input(company_name, possible_sites)
        print(f"Best fuzzy match: {best_match[0]} with confidence {best_match[1]}%")
        return best_match[0]
    
    print("No valid website found.")
    return None

# --- Test with a list of companies ---
companies = ["Tesla", "SpaceX", "Harvard University"]
possible_urls = ["spacex.com", "space-x.com", "spacexofficial.com"]

for company in companies:
    print(f"Searching for website of {company}...")
    official_website = get_official_website(company, possible_sites=possible_urls)
    if official_website:
        print(f"Official website for {company}: {official_website}")
    print("\n" + "-"*50 + "\n")
