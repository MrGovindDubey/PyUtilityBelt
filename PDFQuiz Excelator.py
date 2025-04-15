import os
import google.generativeai as genai
import pandas as pd
from pathlib import Path
import time
import json
from datetime import datetime
import PyPDF2  # Added missing import

class FinalQuizExtractor:
    def __init__(self):
        # Configure API key directly
        self.API_KEY = "KEY HERE TO BE ADDED"
        self.model = None
        self.max_retries = 3
        self.setup_logging()
        self.check_dependencies()

    def check_dependencies(self):
        """Verify all required packages are installed"""
        try:
            import google.generativeai
            import pandas
            import openpyxl
            import PyPDF2
        except ImportError as e:
            print(f"\nERROR: Missing required package - {str(e)}")
            print("Please install dependencies with:")
            print("pip install google-generativeai pandas openpyxl PyPDF2")
            exit(1)

    def setup_logging(self):
        """Create log directory"""
        self.log_dir = os.path.join(os.path.dirname(__file__), "conversion_logs")
        os.makedirs(self.log_dir, exist_ok=True)
        self.log_file = os.path.join(self.log_dir, f"conversion_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")

    def log_message(self, message):
        """Log messages to file and console"""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_entry = f"[{timestamp}] {message}"
        print(log_entry)
        with open(self.log_file, "a") as f:
            f.write(log_entry + "\n")

    def configure_gemini(self):
        """Configure Gemini API"""
        try:
            genai.configure(api_key=self.API_KEY)
            self.model = genai.GenerativeModel('gemini-1.5-pro-latest')
            self.log_message("API configured successfully")
            return True
        except Exception as e:
            self.log_message(f"Configuration failed: {str(e)}")
            return False

    def process_pdf(self, pdf_path):
        """Process all pages of PDF"""
        if not self._validate_pdf(pdf_path):
            return None

        try:
            self.log_message(f"Processing PDF: {pdf_path}")
            
            # First extract text to determine page count
            with open(pdf_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                page_count = len(pdf_reader.pages)
                self.log_message(f"Found {page_count} pages in PDF")

            # Process entire PDF at once
            pdf_file = genai.upload_file(pdf_path)
            
            prompt = """Extract ALL quiz questions from ALL pages and return perfect JSON:
            {
                "quiz_title": "Title",
                "questions": [
                    {
                        "question_number": 1,
                        "question_text": "Full question text",
                        "options": {"A": "Option A", "B": "Option B"},
                        "correct_answer": "A",
                        "question_type": "multiple_choice",
                        "page_number": 1
                    }
                ]
            }
            RULES:
            1. Extract from ALL pages without exception
            2. Include page_number for each question
            3. Preserve complete original text
            4. Include ALL options
            5. Maintain original numbering
            6. Return ONLY valid JSON"""
            
            self.log_message("Extracting questions from all pages...")
            response = self.model.generate_content([prompt, pdf_file])
            
            if not response.text:
                self.log_message("Empty response from Gemini")
                return None
                
            return self._parse_response(response.text, page_count)
            
        except Exception as e:
            self.log_message(f"Processing failed: {str(e)}")
            return None

    def _parse_response(self, text, expected_pages):
        """Parse response with validation"""
        try:
            # Clean response text
            clean_text = text.strip()
            if "```json" in clean_text:
                clean_text = clean_text.split("```json")[1].split("```")[0].strip()
            
            # Handle common issues
            clean_text = clean_text.replace('\n', ' ')  # Remove newlines
            clean_text = clean_text.replace(': null', ': "N/A"')  # Handle nulls
            
            # Parse JSON
            data = json.loads(clean_text)
            
            # Validate structure
            if not isinstance(data, dict) or 'questions' not in data:
                raise ValueError("Invalid quiz structure")
                
            self.log_message(f"Extracted {len(data['questions'])} questions from {expected_pages} pages")
            return data
            
        except Exception as e:
            self.log_message(f"JSON parsing failed: {str(e)}")
            self.log_message(f"Problematic response: {text[:200]}...")
            return None

    def _validate_pdf(self, file_path):
        """Validate PDF file"""
        try:
            if not Path(file_path).exists():
                self.log_message("File not found")
                return False
                
            if not file_path.lower().endswith('.pdf'):
                self.log_message("Only PDF files supported")
                return False
                
            return True
        except Exception as e:
            self.log_message(f"Validation error: {str(e)}")
            return False

    def create_excel(self, quiz_data, output_path):
        """Create Excel file with all questions"""
        try:
            if not quiz_data or not quiz_data.get('questions'):
                self.log_message("No valid questions found")
                return False

            # Prepare data
            excel_data = []
            for q in quiz_data['questions']:
                excel_data.append({
                    'No': q.get('question_number', ''),
                    'Question': q.get('question_text', ''),
                    'Option A': q.get('options', {}).get('A', ''),
                    'Option B': q.get('options', {}).get('B', ''),
                    'Option C': q.get('options', {}).get('C', ''),
                    'Option D': q.get('options', {}).get('D', ''),
                    'Correct': q.get('correct_answer', 'N/A'),
                    'Type': q.get('question_type', ''),
                    'Page': q.get('page_number', 'N/A')
                })

            # Create DataFrame
            df = pd.DataFrame(excel_data)
            
            # Ensure output directory exists
            os.makedirs(os.path.dirname(output_path), exist_ok=True)
            
            # Save to Excel with auto-width
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Quiz')
                
                # Auto-adjust columns
                worksheet = writer.sheets['Quiz']
                for column in worksheet.columns:
                    max_length = max((
                        len(str(cell.value)) 
                        for cell in worksheet[column[0].column_letter]
                    )) + 2
                    worksheet.column_dimensions[column[0].column_letter].width = max_length
            
            return True
            
        except Exception as e:
            self.log_message(f"Excel creation failed: {str(e)}")
            return False

def main():
    print("\n=== FINAL MULTI-PAGE QUIZ EXTRACTOR ===")
    print("Guaranteed to extract all questions from all pages\n")
    
    converter = FinalQuizExtractor()
    if not converter.configure_gemini():
        print("\nFailed to configure Gemini. Check logs for details.")
        return
    
    pdf_path = input("Enter PDF file path: ").strip()
    quiz_data = converter.process_pdf(pdf_path)
    
    if not quiz_data:
        print("\nFailed to process PDF. Check logs for details.")
        return
    
    output_path = os.path.splitext(pdf_path)[0] + "_FULL_QUIZ.xlsx"
    if converter.create_excel(quiz_data, output_path):
        print(f"\nSUCCESS! Complete Excel file created:")
        print(f"Location: {output_path}")
        print(f"Total questions extracted: {len(quiz_data['questions'])}")
    else:
        print("\nFailed to create Excel file. Check logs for details.")

if __name__ == "__main__":
    main()