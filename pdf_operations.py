import os
from pdf2docx import Converter
from pypdf import PdfReader, PdfWriter


class PDFOperations:
    """
    A collection of common PDF manipulation operations.

    Supports selecting PDF files, extracting pages, merging files,
    deleting pages, inserting pages from one PDF into another,
    rotating pages, and converting PDF content to Word (.docx) files.

    Page numbers passed to any method in this class are 1-based
    unless otherwise noted.
    """

    def __init__(self):
        """Initialize the operations handler with an empty file selection."""
        self.pdf_paths = []

    def select_pdf_files(self, paths):
        """
        Select PDF files by specifying their paths.

        :param paths: A single path or a list of paths to PDF files.
        """
        if isinstance(paths, str):
            paths = [paths]  # Convert to list if a single path is provided
        for path in paths:
            if os.path.isfile(path) and path.lower().endswith('.pdf'):
                self.pdf_paths.append(path)
            else:
                print(f"Invalid PDF file: {path}")

    def get_selected_files(self):
        """
        Get the list of selected PDF file paths.

        :return: List of selected PDF file paths.
        """
        return self.pdf_paths

    def extract_pdf_pages(self, pdf_path, page_specifications, output_filename):
        """
        Extract specific pages from a PDF file.

        :param pdf_path: Path to the PDF file.
        :param page_specifications: A single page number, a range (e.g., "1-5"), or a list of pages.
        :param output_filename: Filename to save the extracted pages.
        :return: A PdfWriter object containing the extracted pages.
        """
        if pdf_path not in self.pdf_paths:
            print(f"PDF file {pdf_path} is not selected.")
            return None
        extracted_writer = PdfWriter()
        try:
            with open(pdf_path, 'rb') as file:
                reader = PdfReader(file)
                total_pages = len(reader.pages)

                # Process the page specifications
                pages_to_extract = set()  # Use a set to avoid duplicates

                if isinstance(page_specifications, int):
                    # Single page
                    if 1 <= page_specifications <= total_pages:
                        pages_to_extract.add(page_specifications - 1)
                    else:
                        print(f"Page number {page_specifications} is out of range for {pdf_path}.")
                elif isinstance(page_specifications, str):
                    # Range of pages
                    if '-' in page_specifications:
                        start, end = map(int, page_specifications.split('-'))
                        for page in range(start - 1, end):  # Convert to 0-indexed
                            if 0 <= page < total_pages:
                                pages_to_extract.add(page)
                            else:
                                print(f"Page number {page + 1} is out of range for {pdf_path}.")
                    else:
                        # Single page as string (1-based)
                        page_number = int(page_specifications)
                        if 1 <= page_number <= total_pages:
                            pages_to_extract.add(page_number - 1)
                        else:
                            print(f"Page number {page_number} is out of range for {pdf_path}.")
                elif isinstance(page_specifications, list):
                    # List of pages
                    for page_number in page_specifications:
                        if 1 <= page_number <= total_pages:
                            pages_to_extract.add(page_number - 1)
                        else:
                            print(f"Page number {page_number} is out of range for {pdf_path}.")

                # Add the valid pages to the writer
                for page in sorted(pages_to_extract):
                    extracted_writer.add_page(reader.pages[page])

            # Save the extracted pages to a new PDF file if an output filename is provided
            if output_filename:
                with open(output_filename, 'wb') as output_file:
                    extracted_writer.write(output_file)
                print(f"Extracted pages saved to {output_filename}")

            return extracted_writer
        except Exception as e:
            print(f"An error occurred while extracting pages: {e}")
            return None

    def merge_pdf_files(self, output_filename):
        """
        Merge all selected PDF files into a single PDF file.

        :param output_filename: The filename of the merged PDF.
        """
        if not self.pdf_paths:
            print("No PDF files selected for merging.")
            return

        merged_writer = PdfWriter()
        try:
            for pdf_path in self.pdf_paths:
                with open(pdf_path, 'rb') as file:
                    reader = PdfReader(file)
                    for page in reader.pages:
                        merged_writer.add_page(page)

            # Write once, after all selected files have been added
            with open(output_filename, 'wb') as output_file:
                merged_writer.write(output_file)
            print(f"Merged PDF saved as {output_filename}")
        except Exception as e:
            print(f"An error occurred while merging PDF files: {e}")

    def delete_pdf_pages(self, pdf_path, page_specifications):
        """
        Delete specific pages from a PDF file and overwrite the original file.

        :param pdf_path: Path to the PDF file.
        :param page_specifications: A single page number, a range (e.g., "1-5"), or a list of pages (1-based indexing).
        """
        if pdf_path not in self.pdf_paths:
            print(f"PDF file {pdf_path} is not selected.")
            return None

        modified_writer = PdfWriter()
        try:
            with open(pdf_path, 'rb') as file:
                reader = PdfReader(file)
                total_pages = len(reader.pages)

                # Process the page specifications
                pages_to_delete = set()  # Use a set to avoid duplicates

                if isinstance(page_specifications, int):
                    # Single page (1-based)
                    if 1 <= page_specifications <= total_pages:
                        pages_to_delete.add(page_specifications - 1)  # Convert to 0-indexed
                    else:
                        print(f"Page number {page_specifications} is out of range for {pdf_path}.")
                elif isinstance(page_specifications, str):
                    # Range of pages (1-based)
                    if '-' in page_specifications:
                        start, end = map(int, page_specifications.split('-'))
                        for page in range(start - 1, end):  # Convert start to 0-indexed
                            if 0 <= page < total_pages:
                                pages_to_delete.add(page)
                            else:
                                print(f"Page number {page + 1} is out of range for {pdf_path}.")
                    else:
                        # Single page as string (1-based)
                        page_number = int(page_specifications)
                        if 1 <= page_number <= total_pages:
                            pages_to_delete.add(page_number - 1)  # Convert to 0-indexed
                        else:
                            print(f"Page number {page_number} is out of range for {pdf_path}.")
                elif isinstance(page_specifications, list):
                    # List of pages (1-based)
                    for page_number in page_specifications:
                        if 1 <= page_number <= total_pages:
                            pages_to_delete.add(page_number - 1)  # Convert to 0-indexed
                        else:
                            print(f"Page number {page_number} is out of range for {pdf_path}.")

                # Add pages that are not deleted to the writer
                for page in range(total_pages):
                    if page not in pages_to_delete:
                        modified_writer.add_page(reader.pages[page])

            # Overwrite the original PDF file with the modified content
            with open(pdf_path, 'wb') as output_file:
                modified_writer.write(output_file)
            print(f"Original PDF file {pdf_path} has been modified.")
            return modified_writer
        except Exception as e:
            print(f"An error occurred while deleting pages: {e}")
            return None

    def insert_into_pdf(self, target_pdf_path, source_pdf_path, page_number, insert_position):
        """
        Insert a page from one PDF into another PDF at a specified position.

        :param target_pdf_path: Path to the target PDF file where the page will be inserted.
        :param source_pdf_path: Path to the source PDF file from which the page will be taken.
        :param page_number: The page number (1-based) from the source PDF to insert.
        :param insert_position: The position (1-based) in the target PDF where the page will be inserted.
        """
        if target_pdf_path not in self.pdf_paths:
            print(f"Target PDF file {target_pdf_path} is not selected.")
            return None
        if source_pdf_path not in self.pdf_paths:
            print(f"Source PDF file {source_pdf_path} is not selected.")
            return None

        try:
            with open(target_pdf_path, 'rb') as target_file, open(source_pdf_path, 'rb') as source_file:
                target_reader = PdfReader(target_file)
                source_reader = PdfReader(source_file)

                total_target_pages = len(target_reader.pages)
                total_source_pages = len(source_reader.pages)

                # Convert to 0-based indexing
                if 1 <= page_number <= total_source_pages and 1 <= insert_position <= total_target_pages + 1:
                    page_to_insert = source_reader.pages[page_number - 1]
                    modified_writer = PdfWriter()

                    # Add pages from the target PDF up to the insert position
                    for i in range(insert_position - 1):
                        modified_writer.add_page(target_reader.pages[i])

                    # Insert the page from the source PDF
                    modified_writer.add_page(page_to_insert)

                    # Add the remaining pages from the target PDF
                    for i in range(insert_position - 1, total_target_pages):
                        modified_writer.add_page(target_reader.pages[i])

                    # Overwrite the target PDF with the modified content
                    with open(target_pdf_path, 'wb') as output_file:
                        modified_writer.write(output_file)
                    print(f"Inserted page {page_number} from {source_pdf_path} into {target_pdf_path} at position {insert_position}.")
                else:
                    print("Invalid page number or insert position.")
        except Exception as e:
            print(f"An error occurred while inserting pages: {e}")
            return None

    def rotate_pdf_pages(self, pdf_path, page_specifications, angle):
        """
        Rotate specific pages in a PDF file by a given angle.

        :param pdf_path: Path to the PDF file.
        :param page_specifications: A single page number, a range (e.g., "1-5"), or a list of pages (1-based indexing).
        :param angle: The angle to rotate the pages (90, 180, or 270 degrees).
        """
        if pdf_path not in self.pdf_paths:
            print(f"PDF file {pdf_path} is not selected.")
            return None

        if angle not in [90, 180, 270]:
            print("Invalid angle. Please use 90, 180, or 270 degrees.")
            return None

        modified_writer = PdfWriter()
        try:
            with open(pdf_path, 'rb') as file:
                reader = PdfReader(file)
                total_pages = len(reader.pages)

                # Process the page specifications
                pages_to_rotate = set()

                if isinstance(page_specifications, int):
                    if 1 <= page_specifications <= total_pages:
                        pages_to_rotate.add(page_specifications - 1)  # Convert to 0-indexed
                    else:
                        print(f"Page number {page_specifications} is out of range for {pdf_path}.")
                elif isinstance(page_specifications, str):
                    if '-' in page_specifications:
                        start, end = map(int, page_specifications.split('-'))
                        for page in range(start - 1, end):  # Convert to 0-indexed
                            if 0 <= page < total_pages:
                                pages_to_rotate.add(page)
                            else:
                                print(f"Page number {page + 1} is out of range for {pdf_path}.")
                    else:
                        page_number = int(page_specifications)
                        if 1 <= page_number <= total_pages:
                            pages_to_rotate.add(page_number - 1)  # Convert to 0-indexed
                        else:
                            print(f"Page number {page_number} is out of range for {pdf_path}.")
                elif isinstance(page_specifications, list):
                    for page_number in page_specifications:
                        if 1 <= page_number <= total_pages:
                            pages_to_rotate.add(page_number - 1)  # Convert to 0-indexed
                        else:
                            print(f"Page number {page_number} is out of range for {pdf_path}.")

                # Add pages to the writer, rotating specified pages
                for page in range(total_pages):
                    if page in pages_to_rotate:
                        reader.pages[page].rotate(angle)  # Rotate the page (rotate_clockwise is deprecated)
                    modified_writer.add_page(reader.pages[page])

            # Overwrite the original PDF file with the modified content
            with open(pdf_path, 'wb') as output_file:
                modified_writer.write(output_file)
            print(f"Rotated specified pages in {pdf_path} by {angle} degrees.")
            return modified_writer
        except Exception as e:
            print(f"An error occurred while rotating pages: {e}")
            return None

    def convert_pdf_to_word(self, pdf_path, docx_path, page_specifications=None):
        """
        Convert a PDF file (or specific pages of it) to a Word document.

        :param pdf_path: Path to the PDF file to convert.
        :param docx_path: Path to save the converted Word document.
        :param page_specifications: A single page number, a range (e.g., "1-5"), or a list of pages
            (1-based indexing). If None, the entire document is converted.
        """
        if pdf_path not in self.pdf_paths:
            print(f"PDF file {pdf_path} is not selected.")
            return None

        try:
            cv = Converter(pdf_path)

            if page_specifications is None:
                # Convert the entire document
                cv.convert(docx_path)  # All pages by default
                print(f"Converted entire PDF to {docx_path}.")
            else:
                pages_to_convert = []

                if isinstance(page_specifications, int):
                    if 1 <= page_specifications <= len(cv.pages):
                        pages_to_convert.append(page_specifications - 1)  # Convert to 0-indexed
                    else:
                        print(f"Page number {page_specifications} is out of range for {pdf_path}.")
                elif isinstance(page_specifications, str):
                    if '-' in page_specifications:
                        start, end = map(int, page_specifications.split('-'))
                        for page in range(start - 1, end):  # Convert to 0-indexed
                            if 0 <= page < len(cv.pages):
                                pages_to_convert.append(page)
                            else:
                                print(f"Page number {page + 1} is out of range for {pdf_path}.")
                    else:
                        page_number = int(page_specifications)
                        if 1 <= page_number <= len(cv.pages):
                            pages_to_convert.append(page_number - 1)  # Convert to 0-indexed
                        else:
                            print(f"Page number {page_number} is out of range for {pdf_path}.")
                elif isinstance(page_specifications, list):
                    for page_number in page_specifications:
                        if 1 <= page_number <= len(cv.pages):
                            pages_to_convert.append(page_number - 1)  # Convert to 0-indexed
                        else:
                            print(f"Page number {page_number} is out of range for {pdf_path}.")

                # Convert specified pages
                if pages_to_convert:
                    cv.convert(docx_path, pages=pages_to_convert)
                    print(f"Converted specified pages {page_specifications} from {pdf_path} to {docx_path}.")

            cv.close()
        except Exception as e:
            print(f"An error occurred while converting PDF to Word: {e}")
            return None


if __name__ == "__main__":
    pdf_operations = PDFOperations()
    pdf_operations.select_pdf_files(['ISC_Notes.pdf'])

    # ----------------------------------------------------------------------
    # EXTRACT
    # ----------------------------------------------------------------------

    # Extract a single page
    # extracted_writer_single = pdf_operations.extract_pdf_pages('Maths.pdf', 2, "extracted.pdf")

    # Extract a range of pages and save
    # extracted_writer_single = pdf_operations.extract_pdf_pages('Maths.pdf', "1-5", "extracted.pdf")

    # Extract non-contiguous pages
    # extracted_writer_single = pdf_operations.extract_pdf_pages('Maths.pdf', [1, 3, 5], "extracted.pdf")

    # ----------------------------------------------------------------------
    # MERGE
    # ----------------------------------------------------------------------
    # Merge the selected PDF files into a single file
    # pdf_operations.merge_pdf_files('ISC_Notes.pdf')

    # ----------------------------------------------------------------------
    # DELETE
    # ----------------------------------------------------------------------

    # Delete pages from the original file (1-based)
    # pdf_operations.delete_pdf_pages('merged_output.pdf', [1, 3, 5])

    # Delete a single page from the original file (1-based)
    # pdf_operations.delete_pdf_pages('file1.pdf', 1)

    # Delete a range of pages from the original file (1-based)
    # pdf_operations.delete_pdf_pages('file1.pdf', "1-5")

    # ----------------------------------------------------------------------
    # INSERT
    # ----------------------------------------------------------------------

    # Insert a page from file2.pdf into merged_output.pdf
    # pdf_operations.insert_into_pdf('merged_output.pdf', 'file2.pdf', 1, 2)  # Insert page 1 from file2.pdf at position 2 in merged_output.pdf

    # ----------------------------------------------------------------------
    # ROTATE
    # ----------------------------------------------------------------------

    # Rotate page 1 in merged_output.pdf by 90 degrees
    # pdf_operations.rotate_pdf_pages('merged_output.pdf', 1, 90)

    # Rotate pages 2-3 in merged_output.pdf by 180 degrees
    # pdf_operations.rotate_pdf_pages('merged_output.pdf', "2-3", 180)

    # Rotate non-contiguous pages in merged_output.pdf
    # pdf_operations.rotate_pdf_pages('merged_output.pdf', [1, 3], 270)

    # ----------------------------------------------------------------------
    # Convert PDF document to Word
    # ----------------------------------------------------------------------
    # Convert the entire PDF to Word
    pdf_operations.convert_pdf_to_word('ISC_Notes.pdf', 'file1.docx')

    # Convert specific pages (1-based indexing)
    # pdf_operations.convert_pdf_to_word('file1.pdf', 'file1_partial.docx', page_specifications="1-3")

    # Convert non-contiguous pages
    # pdf_operations.convert_pdf_to_word('file1.pdf', 'file1_selected.docx', page_specifications=[1, 3, 5])
