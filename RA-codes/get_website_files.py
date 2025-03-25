import requests
from bs4 import BeautifulSoup
import os
from urllib.parse import urljoin

base_url = "https://opsportal.spp.org/Studies/Gen"
download_folder = "downloaded studies"

def download_file(url):
    filename = url.split("/")[-1]
    filepath = os.path.join(download_folder, filename)
    response = requests.get(url)
    with open(filepath, "wb") as file:
        file.write(response.content)
    print(f"Downloaded: {filename}")

# Create the download folder if it doesn't exist
os.makedirs(download_folder, exist_ok=True)

# Send a GET request to the base URL
response = requests.get(base_url)

# Create a BeautifulSoup object from the response text
soup = BeautifulSoup(response.text, "html.parser")

# Find all links on the webpage that begin with "/Studies/"
links = soup.find_all("a", href=lambda href: href and href.startswith("/Studies/"))

# Process each link
for link in links:
    # Construct the absolute URL by combining the base URL and the link's href attribute
    link_url = urljoin(base_url, link["href"])

    # Send a GET request to the link URL
    link_response = requests.get(link_url)

    # Create a BeautifulSoup object from the link response text
    link_soup = BeautifulSoup(link_response.text, "html.parser")

    # Find all download links on the subsequent webpage
    download_links = link_soup.find_all("a", href=True)

    # Process each download link
    for download_link in download_links:
        # Check if the download link leads to a file
        if download_link["href"].endswith((".pdf", ".xlsx", ".csv")):
            # Construct the absolute URL for the file download
            download_url = urljoin(link_url, download_link["href"])

            # Download the file
            download_file(download_url)
