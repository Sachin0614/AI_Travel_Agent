import pandas as pd
import streamlit as st
import logging

logger = logging.getLogger(__name__)


def clean_currency_and_convert(value):
    """
    Clean and convert currency values to float
    Handles various currency formats: $1200, USD 1200, "1,200", etc.
    """
    try:
        # Check if the value is missing or Not-a-Number (NaN)
        if pd.isna(value):
            return 0.0
        
        # Convert the value to a string data type
        string_value = str(value).strip()
        
        # Remove currency symbols and commas
        for symbol in ['$', 'USD', '€', '₹', '£', ',', ' ']:
            string_value = string_value.replace(symbol, '')
        
        string_value = string_value.strip()
        
        # Convert to float
        if string_value:
            return float(string_value)
        else:
            return 0.0
            
    except (ValueError, TypeError) as e:
        logger.warning(f"Error converting currency '{value}': {e}")
        return 0.0


@st.cache_data
def load_and_prepare_dataset(file_path):
    """
    Load and prepare travel dataset with comprehensive error handling
    
    Args:
        file_path: Path to CSV file
        
    Returns:
        Prepared DataFrame or None if error
    """
    try:
        # Read the CSV file
        dataframe = pd.read_csv(file_path)
        logger.info(f"Loaded dataset with shape: {dataframe.shape}")
        
        # Standardize column names
        column_mapping = {
            'Traveler gender': 'Gender',
            'Traveler Gender': 'Gender',
            'Accommodation cost': 'Hotel_Price',
            'Accommodation type': 'Hotel_Name',
            'Traveler name': 'Name',
            'Traveler age': 'Age',
        }
        
        dataframe.rename(columns=column_mapping, inplace=True)
        
        # Clean the price column
        if 'Hotel_Price' in dataframe.columns:
            dataframe['Cleaned_Price'] = dataframe['Hotel_Price'].apply(clean_currency_and_convert)
        else:
            dataframe['Cleaned_Price'] = 0.0
            logger.warning("Hotel_Price column not found, using default 0")
        
        # Fill missing values
        dataframe.fillna({
            'Gender': 'Unknown',
            'Hotel_Name': 'Not Specified',
            'Cleaned_Price': 0.0,
        }, inplace=True)
        
        logger.info("Dataset preparation completed successfully")
        return dataframe
        
    except FileNotFoundError:
        logger.error(f"File not found: {file_path}")
        return None
    except Exception as e:
        logger.error(f"Error loading dataset: {e}")
        return None


def get_dataset_summary(dataframe):
    """Get summary statistics of the dataset"""
    if dataframe is None or dataframe.empty:
        return {
            'total_records': 0,
            'total_travelers': 0,
            'unique_destinations': 0,
            'avg_price': 0,
            'max_price': 0,
            'min_price': 0,
        }
    
    return {
        'total_records': len(dataframe),
        'total_travelers': len(dataframe),
        'unique_destinations': dataframe['Destination'].nunique() if 'Destination' in dataframe.columns else 0,
        'avg_price': dataframe['Cleaned_Price'].mean() if 'Cleaned_Price' in dataframe.columns else 0,
        'max_price': dataframe['Cleaned_Price'].max() if 'Cleaned_Price' in dataframe.columns else 0,
        'min_price': dataframe['Cleaned_Price'].min() if 'Cleaned_Price' in dataframe.columns else 0,
    }