import os
import glob
from pathlib import Path
from pptx import Presentation

def extract_links_sequence(file_path):

    prs = Presentation(file_path)

# Initialize variables to keep track of sequence and slide number
    sequence = 1
    slide_number = 1

# Initialize an empty list to store the HTML/Markdown content
    content = []

# Iterate through slides in the presentation
    for slide in prs.slides:
        #sequence = slide.timeline.main_sequence
        for shape in slide.shapes:
        # Check if the shape is a text box
            if shape.has_text_frame:
                text_frame = shape.text_frame
                if text_frame.paragraphs:
                # Extract text from the text box
                    text = ''
                    for paragraph in text_frame.paragraphs:
                        for run in paragraph.runs:
                            text += run.text.strip()
                            
                    # Extract hyperlinks (if any)
                    hyperlinks = []
                    for paragraph in text_frame.paragraphs:
                        for run in paragraph.runs:
                            if run.hyperlink._hlinkClick:
                                for hyperlink in run.hyperlink._hlinkClick:
                                    hyperlinks.append(hyperlink.address)

                    if text:
                        # Generate HTML/Markdown for the text box
                        content.append(f"<div v-click='{sequence}'>{text}")
                        for i, hyperlink in enumerate(hyperlinks):
                            content.append(f"<a href='{hyperlink}' v-click='{sequence + i + 1}'></a>")
                        content.append("</div>")

                # Increment the sequence
                    sequence += len(hyperlinks) + 1

    # Increment slide number
        slide_number += 1

    # Convert the content list to a single string
        html_output = '\n'.join(content)
    return html_output

def get_pptx_file():
    script_directory = os.path.dirname(os.path.abspath(__file__))
   
    #expect hsu-pptx or pptx folder to be in the same directory as this script
    path = Path(script_directory).parent / 'hsu-pptx'
    if not path.is_dir():
        path = Path(script_directory).parent / 'pptx'    
    if not path.is_dir():
        print(f"Error. Files not found in: {Path(script_directory).parent}\n" 
              f"Add a folder named 'pptx' to the same directory as this script and add .pptx files to it.")
        exit()
    # Use glob to filter and sort .pptx files
    file_list = sorted(glob.glob(os.path.join(path, '*.pptx')))

    # Print the list of files to the console
    if file_list:
        print(f"Extract from: ${path}")
        for index, file in enumerate(file_list):
            print(f"{index + 1}. {Path(file).name}")

   # Ask the user to select a file
    while True:
        try:
            selection = int(input("Enter the number of the file you want to extract text from (0 to exit): "))
            
            # Check if the selection is valid
            if 0 <= selection <= len(file_list):
                if selection == 0:
                    print("Exiting...")
                    exit()
                else:
                    selected_file = file_list[selection - 1]
                    print(f"Selected: {selected_file}")
                    return selected_file

            else:
                print("Invalid selection. Please enter a valid number.")
        except ValueError:
            print("Invalid input. Please enter a valid number.")

def main():
    ## data = read_pptx_list()
    file_name = get_pptx_file()
    modified_data = extract_links_sequence(file_name)
    print(modified_data)


if __name__ == "__main__":
    main()