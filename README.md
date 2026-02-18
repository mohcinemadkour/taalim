# Project Documentation

## Overview
This project is a Streamlit-based web application for visualizing and analyzing student statistics, with support for Arabic (RTL) layouts and interactive dashboards. It also includes data exploration scripts for working with Excel files.

## Main Components

### 1. app.py
- **Purpose:** Main Streamlit app for displaying student statistics and dashboards.
- **Key Features:**
  - RTL (Right-to-Left) styling for Arabic language support.
  - Uses Plotly for interactive charts and graphs.
  - Handles file uploads and generates PowerPoint presentations.
  - Customizes Streamlit UI with CSS for RTL and Arabic fonts.

### 2. explore_data.py
- **Purpose:** Data exploration and preprocessing script for Excel files.
- **Key Features:**
  - Reads Excel files to identify the correct header row.
  - Loads and displays data for inspection.
  - Prints column names and data types for further analysis.

## How to Run

### Locally
1. Install dependencies:
   ```sh
   pip install -r requirements.txt
   ```
2. Run the Streamlit app:
   ```sh
   streamlit run app.py
   ```

### With Docker
1. Build the Docker image:
   ```sh
   docker build -t taalim-app .
   ```
2. Run the container:
   ```sh
   docker run -p 10000:10000 taalim-app
   ```
3. Access the app at: http://localhost:10000

## Repository Structure
- `app.py` - Main Streamlit application
- `explore_data.py` - Data exploration script
- `requirements.txt` - Python dependencies
- `Dockerfile` - Containerization instructions
- `test.ipynb` - Jupyter notebook for experiments

## Customization
- Update `app.py` to change dashboard logic or add new features.
- Use `explore_data.py` to preprocess and inspect new datasets before uploading to the app.

## License
This project is licensed under the MIT License.
