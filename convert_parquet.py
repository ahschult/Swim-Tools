import os
import pandas as pd
import pyarrow as pa
import pyarrow.parquet as pq
from datetime import datetime
import xlrd


def clean_dataframe_for_parquet(df):
    df = df.copy()
    
    new_columns = []
    unnamed_count = 0
    for col in df.columns:
        if 'Unnamed:' in str(col):
            if df[col].isna().all() or (df[col].isna().sum() / len(df) > 0.95):
                new_columns.append(None)
            else:
                unnamed_count += 1
                new_columns.append(f'Column_{unnamed_count}')
        else:
            new_columns.append(col)
    
    cols_to_keep = [i for i, col in enumerate(new_columns) if col is not None]
    df = df.iloc[:, cols_to_keep]
    df.columns = [new_columns[i] for i in cols_to_keep]
    
    for col in df.columns:
        if df[col].dtype == 'object':
            df[col] = df[col].astype(str)
            df[col] = df[col].replace(['nan', 'None'], None)
    
    return df


def convert_excel_to_parquet(input_file: str, output_file: str = None, sheet_name: int = 0):
    try:
        if output_file is None:
            base_name = os.path.splitext(input_file)[0]
            output_file = f"{base_name}.parquet"
        
        print(f"\n{'='*80}")
        print(f"Converting Excel to Parquet")
        print(f"{'='*80}")
        print(f"Input file:  {input_file}")
        print(f"Output file: {output_file}")
        print(f"{'='*80}\n")
        
        print("Reading Excel file...")
        _, ext = os.path.splitext(input_file)
        
        if ext.lower() == '.xls':
            df = pd.read_excel(input_file, sheet_name=sheet_name, engine='xlrd')
        else:
            df = pd.read_excel(input_file, sheet_name=sheet_name, engine='openpyxl')
        
        print(f"✓ Successfully read Excel file")
        print(f"  Rows: {len(df)}")
        print(f"  Columns: {len(df.columns)}")
        print(f"\nOriginal column data types:")
        print(df.dtypes)
        
        print(f"\nCleaning data for Parquet compatibility...")
        df = clean_dataframe_for_parquet(df)
        
        print(f"✓ Data cleaned")
        print(f"  Final columns: {len(df.columns)}")
        print(f"\nFinal column data types:")
        print(df.dtypes)
        
        print(f"\nWriting to Parquet format...")
        df.to_parquet(output_file, engine='pyarrow', compression='snappy', index=False)
        
        print(f"✓ Successfully created Parquet file: {output_file}")
        
        input_size = os.path.getsize(input_file) / 1024  # KB
        output_size = os.path.getsize(output_file) / 1024  # KB
        compression_ratio = (1 - output_size / input_size) * 100
        
        print(f"\nFile Size Comparison:")
        print(f"  Excel file:   {input_size:,.2f} KB")
        print(f"  Parquet file: {output_size:,.2f} KB")
        print(f"  Compression:  {compression_ratio:.1f}% smaller")
        print(f"\n{'='*80}\n")
        
        return output_file
        
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()
        return None


def convert_excel_to_parquet_advanced(input_file: str, output_file: str = None, 
                                     sheet_name: int = 0, 
                                     date_columns: list = None,
                                     compression: str = 'snappy',
                                     clean_data: bool = True):
    try:
        if output_file is None:
            base_name = os.path.splitext(input_file)[0]
            output_file = f"{base_name}.parquet"
        
        print(f"\n{'='*80}")
        print(f"Converting Excel to Parquet (Advanced Mode)")
        print(f"{'='*80}")
        print(f"Input file:  {input_file}")
        print(f"Output file: {output_file}")
        print(f"Compression: {compression}")
        print(f"{'='*80}\n")
        
        print("Reading Excel file...")
        _, ext = os.path.splitext(input_file)
        
        read_kwargs = {
            'sheet_name': sheet_name,
            'engine': 'xlrd' if ext.lower() == '.xls' else 'openpyxl'
        }
        
        if date_columns:
            read_kwargs['parse_dates'] = date_columns
            print(f"Parsing date columns: {date_columns}")
        
        df = pd.read_excel(input_file, **read_kwargs)
        
        print(f"✓ Successfully read Excel file")
        print(f"  Shape: {df.shape[0]} rows × {df.shape[1]} columns")
        
        print(f"\nOriginal data types:")
        for col, dtype in df.dtypes.items():
            print(f"  {col:<30} {str(dtype):<15}")
        
        if clean_data:
            print(f"\nCleaning data for Parquet compatibility...")
            df = clean_dataframe_for_parquet(df)
            print(f"✓ Data cleaned")
            print(f"  Final columns: {len(df.columns)}")
            print(f"\nFinal data types:")
            for col, dtype in df.dtypes.items():
                print(f"  {col:<30} {str(dtype):<15}")
        
        print(f"\nWriting to Parquet with {compression} compression...")
        
        df.to_parquet(
            output_file, 
            engine='pyarrow', 
            compression=compression, 
            index=False
        )
        
        print(f"✓ Successfully created Parquet file")
        
        print(f"\nVerifying Parquet file...")
        df_verify = pd.read_parquet(output_file)
        print(f"  Verified rows: {len(df_verify)}")
        print(f"  Verified columns: {len(df_verify.columns)}")
        
        input_size = os.path.getsize(input_file) / 1024
        output_size = os.path.getsize(output_file) / 1024
        compression_ratio = (1 - output_size / input_size) * 100
        
        print(f"\nFile Statistics:")
        print(f"  Original size:  {input_size:,.2f} KB")
        print(f"  Parquet size:   {output_size:,.2f} KB")
        print(f"  Space saved:    {compression_ratio:.1f}%")
        print(f"\n{'='*80}\n")
        
        return output_file
        
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()
        return None


def batch_convert_excel_to_parquet(input_dir: str, output_dir: str = None, 
                                   file_pattern: str = "*.xls*"):
    import glob
    
    if output_dir is None:
        output_dir = input_dir
    
    os.makedirs(output_dir, exist_ok=True)
    
    search_pattern = os.path.join(input_dir, file_pattern)
    excel_files = glob.glob(search_pattern)
    
    print(f"\nFound {len(excel_files)} Excel file(s) to convert")
    print(f"{'='*80}\n")
    
    converted_files = []
    
    for excel_file in excel_files:
        filename = os.path.basename(excel_file)
        base_name = os.path.splitext(filename)[0]
        output_file = os.path.join(output_dir, f"{base_name}.parquet")
        
        print(f"Converting: {filename}")
        result = convert_excel_to_parquet(excel_file, output_file)
        
        if result:
            converted_files.append(result)
    
    print(f"\n{'='*80}")
    print(f"Batch conversion complete: {len(converted_files)}/{len(excel_files)} files converted")
    print(f"{'='*80}\n")
    
    return converted_files


if __name__ == "__main__":
    print("Example 1: Simple Conversion")
    print("-" * 80)
    input_file = "data/your_file.xlsx"  # Change this to your file path
    convert_excel_to_parquet(input_file)
    
    print("\n\nExample 2: Advanced Conversion with Date Columns")
    print("-" * 80)
    print("\n\nExample 3: Batch Conversion")
    print("-" * 80)
    