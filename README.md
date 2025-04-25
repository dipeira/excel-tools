# Excel Tools

A web application for processing Excel files with various tools.

## Features

### 1. Compare Excel Files
- Compare two Excel files based on selected columns
- Find differences between files
- Download a report of the differences

### 2. Join Excel Files
- Join two Excel files based on a common column
- Combine all columns from both files
- Download the joined file

### 3. Filter Excel File
- Filter an Excel file based on a specific value in a selected column
- Extract rows matching the filter criteria
- Download the filtered file

### 4. Split by Column
- Split an Excel file into multiple files based on unique values in a selected column
- Each unique value gets its own Excel file
- All split files are packaged into a zip file for easy download

## Technical Details

### Dependencies
- Python 3.x
- Flask
- Pandas
- Openpyxl

### Installation
1. Clone the repository
2. Create a virtual environment
3. Install dependencies:
   ```
   pip install -r requirements.txt
   ```
4. Run the application:
   ```
   python app.py
   ```

### File Structure
- `app.py`: Main Flask application
- `templates/`: HTML templates
  - `base.html`: Base template with common layout
  - `index.html`: Home page with tool cards
  - `compare.html`: Compare files interface
  - `join.html`: Join files interface
  - `filter.html`: Filter file interface
  - `split.html`: Split file interface
- `uploads/`: Directory for temporary file storage
  - `split_files/`: Directory for split Excel files

### Security Features
- File type validation (.xlsx, .xls)
- Secure filename handling
- Automatic cleanup of temporary files
- Maximum file size limit (16MB)

## Usage

1. Access the web interface at `http://localhost:8000`
2. Choose a tool from the navigation menu
3. Upload Excel file(s) as required
4. Select columns and values as needed
5. Process the file(s)
6. Download the results

## Notes
- All uploaded files are processed in memory and cleaned up after processing
- Split files are automatically zipped for easy download
- The application supports both .xlsx and .xls file formats 