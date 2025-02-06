from bs4 import BeautifulSoup
import os


def extract_text_from_html(file_path):
    """Extracts clean text content from an HTML file."""
    with open(file_path, "r", encoding="utf-8") as file:
        soup = BeautifulSoup(file, "html.parser")

        # Remove script, style, and other unwanted elements
        for tag in soup(["script", "style", "meta", "noscript", "header", "footer", "nav", "aside"]):
            tag.extract()

        # Extract text and clean up spaces        
        text = soup.get_text(separator=" ").strip()
        #print(str(text))
        text = ' '.join(text.split())  # Remove excessive spaces and new lines
        return text

def process_html_files(folder_path):
    """Reads all HTML files from a folder and extracts content."""
    if not os.path.exists(folder_path):
        print("Folder does not exist:", folder_path)
        return
    
    for filename in os.listdir(folder_path):
        if filename.endswith(".html"):  # Process only .html files
            file_path = os.path.join(folder_path, filename)
            print(f"\nProcessing: {filename}")
            
            content = extract_text_from_html(file_path)
            content = content.replace("\ufeff", "")
            
            # Print extracted content (or save it to a file)
            print(content[:500].encode('utf-8', errors='ignore').decode('utf-8'))  # Print first 500 characters for preview
            
            # Optionally, save extracted text to a .txt file
            output_file = os.path.join(folder_path, filename.replace(".html", ".txt"))
            with open(output_file, "w", encoding="utf-8") as out:
                out.write(content)
                print(f"Saved extracted text to: {output_file}")

# Set folder path
folder_path = r"C:\Users\somnath.das\Desktop\Default"  # Replace with your folder path
process_html_files(folder_path)
