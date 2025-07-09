import os
import logging
from pathlib import Path
from .config import MESSAGES, DEFAULT_WIDTH_PERCENT
from .splitter import ExcelSplitter
from .merger import ExcelMerger

class ExcelProcessor:
    def __init__(self):
        self.language = self._select_language()
        self.messages = MESSAGES[self.language]
        self._setup_logging()
        self.splitter = ExcelSplitter(self.messages, self.language)
        self.merger = ExcelMerger(self.messages, self.language)

    def _setup_logging(self):
        """Configure logging based on selected language"""
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            filename='excel_tools.log',
            encoding='utf-8'
        )

    def _select_language(self):
        """Language selection with dual prompt"""
        print("Select language [E-English/R-Русский]:")
        print("Выберите язык [E-Английский/R-Русский]:")
        while True:
            lang = input().upper()
            if lang in ['E', 'R']:
                return lang
            print("Please enter E or R / Пожалуйста введите E или R")

    def _get_width_percent(self):
        """Get width reduction percentage with validation"""
        try:
            percent = float(input(self.messages['width_factor']) or DEFAULT_WIDTH_PERCENT)
            return max(1, min(100, percent))
        except ValueError:
            print(self.messages['invalid_width'])
            return DEFAULT_WIDTH_PERCENT

    def process_files(self):
        print(f"\n{self.messages['welcome']}")
        
        if self.language == 'R':
            action = input("Выберите действие (1-разделить/2-объединить): ")
            if action == '2':
                return self._process_merge()
        
        return self._process_split()

    def _process_split(self):
        """Обработка разделения с запросом на добавление изображения"""
        input_file = input(self.messages['select_input'])
        output_folder = input(self.messages['select_output'])
        card_marker = input(self.messages['card_marker'])
        width_percent = self._get_width_percent()
        
        # Запрос на добавление изображения
        if self.language == 'R':
            add_image = input("Добавить изображение в карточки? (y/n): ").lower() == 'y'
        else:
            add_image = input("Add image to cards? (y/n): ").lower() == 'y'
        
        image_path = None
        if add_image:
            image_path = input("Введите путь к изображению: " if self.language == 'R' 
                             else "Enter image path: ")
            if not self.splitter.set_image(image_path):
                print(self.messages['image_not_found'])
        
        return self.splitter.process_all_cards(input_file, output_folder, 
                                             card_marker, width_percent)

    def _process_merge(self):
        """Handle file merging operation"""
        try:
            input_folder = input("Введите папку с файлами для объединения: ")
            output_file = input("Введите имя итогового файла (например: merged.xlsx): ")

            # Normalize paths for Windows
            input_folder = os.path.normpath(input_folder.strip())
            output_file = os.path.normpath(output_file.strip())

            return self.merger.merge_files(input_folder, output_file)
        except Exception as e:
            print(f"{self.messages['error']}: {str(e)}")
            return False