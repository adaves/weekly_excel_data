import unittest
from datetime import datetime, timedelta
import os
import shutil
from unittest.mock import patch, Mock
import tempfile
import circana_data_script

class TestCircanaDataScript(unittest.TestCase):
    def test_extract_date_dot_format(self):
        """Test extracting date from filenames with format MM.DD.YY"""
        filename = "MULOplus_Circana Weekly Dollar and Unit Consumption Trends_L1 and CYTD Through WE 04.27.25.xlsx"
        expected = (4, 27, 2025)
        self.assertEqual(circana_data_script.extract_date_from_filename(filename), expected)
    
    def test_extract_date_no_dots(self):
        """Test extracting date from filenames with format MMDDYY"""
        filename = "MULOplus_Circana Weekly Dollar and Unit Consumption Trends_L1 and CYTD Through WE 042025.xlsx"
        expected = (4, 20, 2025)
        self.assertEqual(circana_data_script.extract_date_from_filename(filename), expected)
    
    def test_extract_date_fallback(self):
        """Test fallback to previous Sunday when no date in filename"""
        filename = "MULOplus_Circana Weekly Dollar and Unit Consumption.xlsx"
        
        # Mock today's date to a known value (Wednesday, May 15, 2024)
        with patch('circana_data_script.datetime') as mock_datetime:
            mock_datetime.now.return_value = datetime(2024, 5, 15)
            mock_datetime.side_effect = lambda *args, **kw: datetime(*args, **kw)
            
            # Previous Sunday would be May 12, 2024
            expected = (5, 12, 2024)
            self.assertEqual(circana_data_script.extract_date_from_filename(filename), expected)
    
    def test_format_date(self):
        """Test converting date tuple to mm-dd-yyyy format"""
        date_tuple = (4, 27, 2025)
        expected = "04-27-2025"
        self.assertEqual(circana_data_script.format_date(date_tuple), expected)
    
    def test_create_new_filename(self):
        """Test creating new filename with date prefix"""
        original_filename = "MULOplus_Circana Weekly Dollar and Unit Consumption Trends_L1 and CYTD Through WE 04.27.25.xlsx"
        expected = "04-27-2025_MULOplus_Circana Weekly Dollar and Unit Consumption Trends_L1 and CYTD Through WE 04.27.25.xlsx"
        self.assertEqual(circana_data_script.create_new_filename(original_filename), expected)
    
    def test_get_output_path(self):
        """Test creating output path with modified_excel_workbooks directory"""
        filename = "04-27-2025_test.xlsx"
        expected = os.path.join("modified_excel_workbooks", filename)
        self.assertEqual(circana_data_script.get_output_path(filename), expected)

    def test_get_archive_path(self):
        """Test creating archive path with archived_data directory"""
        filename = "test.xlsx"
        expected = os.path.join("archived_data", filename)
        self.assertEqual(circana_data_script.get_archive_path(filename), expected)
    
    def test_find_excel_files(self):
        """Test finding all Excel files in directory"""
        # Create a temp directory and files for testing
        with tempfile.TemporaryDirectory() as tmpdirname:
            # Create test files
            open(os.path.join(tmpdirname, "file1.xlsx"), 'a').close()
            open(os.path.join(tmpdirname, "file2.xls"), 'a').close()
            open(os.path.join(tmpdirname, "file3.txt"), 'a').close()
            
            # Create subdirectories to verify they're not included
            os.makedirs(os.path.join(tmpdirname, "modified_excel_workbooks"))
            os.makedirs(os.path.join(tmpdirname, "archived_data"))
            open(os.path.join(tmpdirname, "modified_excel_workbooks", "file4.xlsx"), 'a').close()
            open(os.path.join(tmpdirname, "archived_data", "file5.xlsx"), 'a').close()
            
            # Test the function
            excel_files = circana_data_script.find_excel_files(tmpdirname)
            
            # Should find 2 Excel files, not including subdirectories
            self.assertEqual(len(excel_files), 2)
            self.assertIn(os.path.join(tmpdirname, "file1.xlsx"), excel_files)
            self.assertIn(os.path.join(tmpdirname, "file2.xls"), excel_files)
    
    def test_process_and_archive(self):
        """Test processing and archiving functionality"""
        # Create temporary directory structure
        with tempfile.TemporaryDirectory() as tmpdirname:
            # Mock Excel file
            mock_file = os.path.join(tmpdirname, "test.xlsx")
            with open(mock_file, 'a') as f:
                f.write("mock content")
            
            # Create output and archive directories
            output_dir = os.path.join(tmpdirname, "modified_excel_workbooks")
            archive_dir = os.path.join(tmpdirname, "archived_data")
            os.makedirs(output_dir)
            os.makedirs(archive_dir)
            
            # Mock the openpyxl functionality
            with patch('circana_data_script.openpyxl.load_workbook') as mock_load:
                mock_wb = Mock()
                mock_sheet = Mock()
                mock_wb.worksheets = [mock_sheet]
                mock_load.return_value = mock_wb
                
                # Mock the extract_date function to return a known date
                with patch('circana_data_script.extract_date_from_filename', return_value=(5, 15, 2024)):
                    # Test with custom paths
                    result = circana_data_script.process_excel_file(
                        mock_file, 
                        output_dir=output_dir,
                        archive_dir=archive_dir
                    )
                    
                    # Check if file was processed correctly
                    expected_output = os.path.join(output_dir, "05-15-2024_test.xlsx")
                    self.assertEqual(result, expected_output)
                    
                    # Check if original was archived
                    expected_archive = os.path.join(archive_dir, "test.xlsx")
                    self.assertTrue(os.path.exists(expected_archive))
                    
                    # Verify original file no longer exists in source directory
                    self.assertFalse(os.path.exists(mock_file))

if __name__ == "__main__":
    unittest.main() 