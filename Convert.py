import os
import shutil
import platform
from PIL import Image
from docx2pdf import convert

def convert_pptx_to_pdf_mac(file_path, pdf_path):
    """Convert PowerPoint to PDF on macOS using AppleScript"""
    import subprocess

    # AppleScript command to convert PPTX to PDF
    script = f'''
    tell application "Microsoft PowerPoint"
        open "{file_path}"
        set pptx_file to active presentation
        save pptx_file in "{pdf_path}" as PDF
        close pptx_file
        quit
    end tell
    '''

    # Execute the AppleScript command
    subprocess.run(['osascript', '-e', script], check=True)

def convert_pptx_to_pdf_windows(file_path, pdf_path):
    """Convert PowerPoint to PDF on Windows using COM"""
    import comtypes.client

    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    presentation = powerpoint.Presentations.Open(os.path.abspath(file_path))
    presentation.SaveAs(os.path.abspath(pdf_path), 32)  # 32 is the PDF format code
    presentation.Close()
    powerpoint.Quit()

def convert_to_pdf(src_dir, dest_dir):
    # Create the destination directory if it doesn't exist
    os.makedirs(dest_dir, exist_ok=True)

    # Determine the operating system
    system = platform.system()

    # Convert src_dir and dest_dir to absolute paths
    src_dir = os.path.abspath(src_dir)
    dest_dir = os.path.abspath(dest_dir)

    # Iterate over files in the source directory
    for filename in os.listdir(src_dir):
        file_path = os.path.join(src_dir, filename)

        # Check if the file is a regular file (not a directory)
        if os.path.isfile(file_path):
            # Get the file extension
            _, extension = os.path.splitext(filename)
            extension = extension.lower()

            try:
                if extension == '.txt':
                    # Convert text file to PDF using reportlab
                    from reportlab.pdfgen import canvas
                    from reportlab.lib.pagesizes import letter

                    pdf_filename = os.path.splitext(filename)[0] + '.pdf'
                    pdf_path = os.path.join(dest_dir, pdf_filename)

                    c = canvas.Canvas(pdf_path, pagesize=letter)
                    with open(file_path, 'r') as txt_file:
                        y = 750  # Starting y position
                        for line in txt_file:
                            if y < 50:  # Check if we need a new page
                                c.showPage()
                                y = 750
                            c.drawString(50, y, line.strip())
                            y -= 15  # Move down for next line
                    c.save()

                elif extension in ['.doc', '.docx']:
                    # Convert Word document to PDF
                    pdf_filename = os.path.splitext(filename)[0] + '.pdf'
                    pdf_path = os.path.join(dest_dir, pdf_filename)
                    convert(file_path, pdf_path)

                elif extension in ['.jpg', '.jpeg', '.png']:
                    # Convert image to PDF
                    pdf_filename = os.path.splitext(filename)[0] + '.pdf'
                    pdf_path = os.path.join(dest_dir, pdf_filename)
                    image = Image.open(file_path)
                    # Convert to RGB if necessary
                    if image.mode not in ('RGB', 'L'):
                        image = image.convert('RGB')
                    image.save(pdf_path, 'PDF')

                elif extension == '.pptx':
                    # Convert PowerPoint to PDF based on OS
                    pdf_filename = os.path.splitext(filename)[0] + '.pdf'
                    pdf_path = os.path.join(dest_dir, pdf_filename)

                    if system == 'Darwin':  # macOS
                        convert_pptx_to_pdf_mac(file_path, pdf_path)
                    elif system == 'Windows':
                        convert_pptx_to_pdf_windows(file_path, pdf_path)
                    else:
                        print(f"PowerPoint conversion not supported on {system}")
                        continue

                elif extension == '.pdf':
                    # Copy files that are already PDFs
                    shutil.copy2(file_path, dest_dir)

                else:
                    # Skip over unsupported file types
                    continue

                print(f"Successfully converted {filename}")

            except Exception as e:
                print(f"Error converting {filename}: {str(e)}")

def main():
    source_directory = input("Enter the source directory: ")
    destination_directory = input("Enter the destination directory: ")

    # Validate directories
    if not os.path.exists(source_directory):
        print("Source directory does not exist!")
        return

    convert_to_pdf(source_directory, destination_directory)
    print("Conversion process completed!")

if __name__ == "__main__":
    main()