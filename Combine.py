import pikepdf
import os
from pathlib import Path


def combine_pdfs(input_dir, output_path):
    # Track files that couldn't be processed
    failed_files = []
    processed_files = []

    try:
        # Create a new PDF to store all pages
        pdf_merger = pikepdf.Pdf.new()

        # Process each PDF in sorted order
        for filename in sorted(os.listdir(input_dir)):
            if filename.lower().endswith('.pdf'):
                input_path = os.path.join(input_dir, filename)
                try:
                    # Open and add each PDF
                    with pikepdf.open(input_path) as pdf:
                        # Add all pages from this PDF
                        pdf_merger.pages.extend(pdf.pages)
                        processed_files.append(filename)
                        print(f"Added: {filename}")

                except Exception as e:
                    failed_files.append((filename, str(e)))
                    print(f"Error processing {filename}: {str(e)}")

        # Save the combined PDF if we processed any files
        if processed_files:
            pdf_merger.save(output_path)
            print(f"\nSuccessfully created combined PDF: {output_path}")
            print(f"Combined {len(processed_files)} files")
        else:
            print("\nNo PDFs were able to be combined")

        # Report any failures
        if failed_files:
            print("\nThe following files could not be combined:")
            for name, reason in failed_files:
                print(f"- {name}: {reason}")

    except Exception as e:
        print(f"Error during combination process: {str(e)}")


def main():
    input_directory = input("Directory of PDFs to be combined: ")
    output_pdf = input_directory + '/Combined.pdf'
    combine_pdfs(input_directory, output_pdf)

if __name__ == '__main__':
    main()