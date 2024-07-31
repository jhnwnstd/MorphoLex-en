import pandas as pd
from pathlib import Path
import warnings
import logging

# Suppress specific openpyxl warnings
warnings.filterwarnings("ignore", category=UserWarning, module='openpyxl')

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(message)s')

class MorphoLEXProcessor:
    def __init__(self, file_path):
        self.file_path = Path(file_path)
        self.engine = 'openpyxl' if self.file_path.suffix == '.xlsx' else 'xlrd'
        self.excel_data = pd.ExcelFile(self.file_path, engine=self.engine)
        self.numbered_sheets = [sheet for sheet in self.excel_data.sheet_names if sheet[0].isdigit()]
        self.data_dict, self.word_index = self._preprocess_data(self.numbered_sheets)
        
    def _preprocess_data(self, sheet_names, usecols=None):
        """Load all numbered sheets into a dictionary and index words for quick lookup."""
        data = {}
        word_index = {}
        
        for sheet in sheet_names:
            try:
                # Specify dtype for 'Word' column to ensure it is read as string
                dtype = {'Word': str}
                df = pd.read_excel(self.file_path, sheet_name=sheet, usecols=usecols, dtype=dtype, engine=self.engine)
                data[sheet] = df
                
                # Vectorized operations for word processing
                df = df.dropna(subset=['Word']).copy()
                df['Word'] = df['Word'].str.strip().str.lower()
                
                for _, row in df.iterrows():
                    word = row['Word']
                    if word not in word_index:
                        word_index[word] = []
                    word_index[word].append((sheet, row))
            except Exception as e:
                logging.warning(f"Error loading sheet {sheet}: {e}")
        
        return data, word_index

    def query_by_word(self, word):
        """Query the preprocessed index by word."""
        word = word.strip().lower()
        return self.word_index.get(word, [])

    def display_results(self, results, fields=None):
        """Display query results in a readable format."""
        for sheet_name, row in results:
            logging.info(f"Results from sheet: {sheet_name}")
            if fields:
                filtered_row = row[fields]
            else:
                filtered_row = row[['MorphoLexSegm']] if 'MorphoLexSegm' in row else row
            logging.info(f"\n{filtered_row.to_string(index=False)}\n")

    def interactive_query(self):
        """Interactive querying tool for MorphoLEX Data."""
        logging.info("Interactive Query Tool for MorphoLEX Data")
        logging.info("Type 'exit.' to quit.\n")
        while True:
            word = input("Enter a word to query: ").strip()
            if word.lower() == 'exit.':
                logging.info("Exiting the query tool. Goodbye!")
                break
            results = self.query_by_word(word)
            if results:
                first_result = results[0][1]
                logging.info("Available fields:")
                logging.info(", ".join(first_result.index))
                fields = input("Enter fields to display (comma-separated), or press Enter for default: ").strip()
                fields = [field.strip() for field in fields.split(",")] if fields else None
                self.display_results(results, fields)
            else:
                logging.info("Word not found in any sheet.")

if __name__ == "__main__":
    processor = MorphoLEXProcessor('main/data/MorphoLEX_en.xls')
    processor.interactive_query()