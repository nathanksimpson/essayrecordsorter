# Essay Data Sorter Application

A simple web application for parsing and filtering essay evaluation data from Excel files or pasted text.

## Features

- **Parse Data**: Import data from Excel files (.xlsx) or paste tab-separated text
- **Filter by Multiple Criteria**: Filter by Class, Essay Type, Evaluator/Method, and Evaluation Status
- **Real-time Filtering**: Results update instantly as you select or deselect filters
- **Checkbox Interface**: Easy-to-use checkboxes automatically generated from your data
- **View Results**: See filtered results in a clean, scrollable table

## How to Use

### Step 1: Input Your Data

You can input data in two ways:

**Option A: Paste Text**
1. Copy your data from Excel or any text source
2. Paste it into the "Paste tab-separated data here" text area
3. Click "Parse Pasted Data"

**Option B: Upload Excel File**
1. Click "Choose File" under "Or upload Excel file:"
2. Select your .xlsx file
3. The file will be parsed automatically

### Step 2: Filter Your Results

After parsing, checkboxes will automatically appear for:
- **Class**: Select one or more classes to filter
- **Essay Type**: Select essay types (e.g., News, Debate)
- **Evaluator/Method**: Select evaluators or methods (empty values shown as "[Empty]")
- **Evaluation Status**: Select evaluation statuses (e.g., Edit(O)-Eval(O), Edit(X)-Eval(X))

**Filter Logic:**
- Selecting multiple checkboxes in the same group shows results that match ANY of those selections
- Selecting checkboxes across different groups shows results that match ALL selected criteria
- If no checkboxes are selected in a group, that group doesn't filter the results

### Step 3: View Results

The results table shows:
- Student Name
- Class
- Code
- Teacher
- Essay Type
- Evaluation
- Evaluator/Method
- Date

Results update in real-time as you change your filter selections.

## Data Format

Your data should be in tab-separated format with 7 columns per row:

1. Student Name (e.g., "안승우S(Eddie A)")
2. Class + Code (e.g., "중1 유마 M^2511")
3. Teacher Name (e.g., "김재진")
4. Essay Type (e.g., "News", "Debate")
5. Evaluation (e.g., "Edit(O)-Eval(O)", "Edit(X)-Eval(X)")
6. Evaluator/Method (e.g., "온라인", "Rachelle", "Nathan", or blank)
7. Date (e.g., "2025-12-08")

## Tips

- **Finding Incomplete Essays**: Check the "Edit(X)-Eval(X)" checkbox in the Evaluation Status filter
- **Clear Filters**: Click "Clear All Filters" to uncheck all filters and see all data
- **Start Over**: Click "Clear All Data" to remove all data and start fresh
- **Multiple Classes**: You can select multiple classes to see essays from different classes at once

## Browser Compatibility

This application works in all modern web browsers:
- Chrome
- Firefox
- Edge
- Safari

No installation required - just open `index.html` in your browser!

## Troubleshooting

**Data won't parse:**
- Make sure your data is tab-separated (copy from Excel usually works)
- Check that each row has 7 columns
- Remove any extra blank rows

**No results showing:**
- Check that at least one checkbox is selected in each filter group you want to use
- Try clicking "Clear All Filters" to see all data
- Verify your data was parsed successfully (check the status message)

**Checkboxes not appearing:**
- Make sure your data was parsed successfully
- Check the status message for any errors
- Try pasting your data again


