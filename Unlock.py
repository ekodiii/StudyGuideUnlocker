import pikepdf
import os
from pathlib import Path

def batch_convert_pdfs(input_dir, output_dir, view_password):
    # Create output directory if it doesn't exist
    os.makedirs(output_dir, exist_ok=True)

    success_count = 0
    failed_files = []

    for filename in os.listdir(input_dir):
        if filename.lower().endswith('.pdf'):
            input_path = os.path.join(input_dir, filename)
            output_path = os.path.join(output_dir, f'printed_{filename}')

            try:
                # Attempt to open and save the PDF
                with pikepdf.open(input_path, password=view_password) as pdf:
                    # Just try to save - if permissions don't allow, it will raise an exception
                    pdf.save(output_path)
                    success_count += 1
                    print(f"Successfully processed: {filename}")

            except pikepdf.PasswordError:
                failed_files.append((filename, "Invalid password"))
            except pikepdf.PdfError as e:
                failed_files.append((filename, f"PDF Error: {str(e)}"))
            except Exception as e:
                failed_files.append((filename, str(e)))
                print(f"Error processing {filename}: {str(e)}")

    # Print summary
    print(f"\nProcessing complete:")
    print(f"Successfully processed: {success_count} files")
    if failed_files:
        print("\nFailed files:")
        for name, reason in failed_files:
            print(f"- {name}: {reason}")


# Usage
input_directory = input("Input directory containing PDF files: ")
output_directory = input("Output directory to put PDF files: ")
view_password = input("Password: ")

batch_convert_pdfs(input_directory, output_directory, view_password)
