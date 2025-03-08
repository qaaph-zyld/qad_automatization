import os
import sys
import pandas as pd
import pyodbc
import glob
import logging
import argparse
from datetime import datetime

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

def parse_arguments():
    """
    Parse command line arguments
    """
    parser = argparse.ArgumentParser(description='Analyze customer demand data with BOM data')
    
    parser.add_argument('--excel-dir', 
                        default=r"C:\Users\ajelacn\AppData\Local\Temp\Shell",
                        help='Directory containing the exported Excel files')
    
    parser.add_argument('--sql-file', 
                        default=r"C:\Users\ajelacn\OneDrive - Adient\Documents\Projects\QAD_automation\sql querries\BOM_PO.sql",
                        help='SQL file with BOM and PO queries')
    
    parser.add_argument('--output', 
                        default=r"C:\Users\ajelacn\OneDrive - Adient\Documents\Projects\QAD_automation\component_demand.xlsx",
                        help='Output file for component demand report')
    
    parser.add_argument('--db-server',
                        default='a265m001',
                        help='Database server name')
    
    parser.add_argument('--db-name',
                        default='QADEE',
                        help='Database name')
    
    parser.add_argument('--verbose', '-v',
                        action='store_true',
                        help='Enable verbose logging')
    
    return parser.parse_args()

def get_latest_excel_file(directory):
    """
    Get the latest Excel file in the specified directory
    """
    try:
        # List all Excel files in the directory
        excel_files = glob.glob(os.path.join(directory, "tmp*.xlsx"))
        
        if not excel_files:
            logger.error(f"No Excel files found in {directory}")
            return None
            
        # Get the latest file based on modification time
        latest_file = max(excel_files, key=os.path.getmtime)
        logger.info(f"Found latest Excel file: {latest_file}")
        return latest_file
    except Exception as e:
        logger.error(f"Error finding latest Excel file: {str(e)}")
        return None

def read_excel_data(file_path):
    """
    Read customer demand data from Excel file
    """
    try:
        logger.info(f"Reading Excel file: {file_path}")
        df = pd.read_excel(file_path)
        
        # Clean up column names (remove whitespace)
        df.columns = [col.strip() for col in df.columns]
        
        # Convert date columns to datetime if needed
        if 'Date' in df.columns:
            df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
        
        logger.info(f"Excel data loaded successfully. {len(df)} rows found.")
        return df
    except Exception as e:
        logger.error(f"Error reading Excel file: {str(e)}")
        return None

def execute_sql_query(sql_file_path, db_server, db_name):
    """
    Execute SQL query from file and return results as DataFrame
    """
    try:
        # Read SQL query from file
        logger.info(f"Reading SQL query from: {sql_file_path}")
        with open(sql_file_path, 'r') as file:
            sql_query = file.read()
        
        logger.info("SQL query loaded successfully")
        
        try:
            # Try to connect to the database and execute the query
            logger.info(f"Attempting to connect to database server: {db_server}, database: {db_name}")
            conn_str = (
                f"DRIVER={{SQL Server}};"
                f"SERVER={db_server};"
                f"DATABASE={db_name};"
                f"Trusted_Connection=yes;"
            )
            
            # Try to connect to the database
            try:
                conn = pyodbc.connect(conn_str)
                logger.info("Connected to database successfully")
                
                # Execute the query
                bom_df = pd.read_sql(sql_query, conn)
                logger.info(f"SQL query executed successfully. {len(bom_df)} rows returned.")
                
                return bom_df
            except Exception as db_error:
                logger.warning(f"Database connection failed: {str(db_error)}")
                logger.warning("Using mock BOM data instead")
                
                # If database connection fails, use mock data
                bom_data = {
                    'Plant': ['2674', '2674', '2798', '2798'],
                    'ps_par': ['PART1', 'PART2', 'PART3', 'PART4'],
                    'ps_comp': ['COMP1', 'COMP2', 'COMP3', 'COMP4'],
                    'pt_desc1': ['Description 1', 'Description 2', 'Description 3', 'Description 4'],
                    'pt_desc2': ['Detail 1', 'Detail 2', 'Detail 3', 'Detail 4'],
                    'ps_qty_per': [2, 3, 1, 4],
                    'po_vend': ['VEND1', 'VEND2', 'VEND3', 'VEND4'],
                    'pt_prod_line': ['LINE1', 'LINE2', 'LINE1', 'LINE2'],
                    'pt_dsgn_grp': ['GROUP1', 'GROUP2', 'GROUP1', 'GROUP2']
                }
                bom_df = pd.DataFrame(bom_data)
                logger.info(f"Mock BOM data created. {len(bom_df)} rows.")
                return bom_df
                
        except Exception as e:
            logger.error(f"Error in database operations: {str(e)}")
            return None
            
    except Exception as e:
        logger.error(f"Error executing SQL query: {str(e)}")
        return None

def analyze_demand_with_bom(demand_df, bom_df, verbose=False):
    """
    Analyze customer demand data with BOM data to calculate component demand
    """
    try:
        logger.info("Analyzing demand data with BOM data...")
        
        # Display demand data summary
        logger.info("\nCustomer Demand Summary:")
        logger.info(f"Total rows: {len(demand_df)}")
        logger.info(f"Columns: {', '.join(demand_df.columns)}")
        
        # Display sample of demand data if verbose
        if verbose:
            logger.info("\nSample of demand data:")
            logger.info(demand_df.head(5))
        
        # Display BOM data summary
        logger.info("\nBOM Data Summary:")
        logger.info(f"Total rows: {len(bom_df)}")
        logger.info(f"Columns: {', '.join(bom_df.columns)}")
        
        # Display sample of BOM data if verbose
        if verbose:
            logger.info("\nSample of BOM data:")
            logger.info(bom_df.head(5))
        
        # Extract part numbers from demand data
        # Assuming 'Item Number' column contains the part numbers
        if 'Item Number' in demand_df.columns:
            # Merge BOM data with demand data
            logger.info("\nMerging BOM data with demand data...")
            
            # Create a copy of the demand dataframe to avoid modifying the original
            demand_copy = demand_df.copy()
            
            # Merge on Item Number (from demand) and ps_par (from BOM)
            merged_data = pd.merge(
                demand_copy,
                bom_df,
                left_on='Item Number',
                right_on='ps_par',
                how='left'
            )
            
            # Check for parts without BOM data
            missing_bom = merged_data[merged_data['ps_par'].isna()]
            if not missing_bom.empty:
                missing_count = len(missing_bom['Item Number'].unique())
                logger.warning(f"\n{missing_count} parts don't have BOM data")
                if verbose:
                    missing_parts = missing_bom['Item Number'].unique()
                    logger.warning(missing_parts[:10].tolist())
                    if len(missing_parts) > 10:
                        logger.warning(f"... and {len(missing_parts) - 10} more")
            
            # Create pivot table with dates as columns
            logger.info("\nCreating pivot table with dates as columns...")
            
            # Format the date column to be used as pivot columns
            merged_data['Date_Formatted'] = merged_data['Date'].dt.strftime('%Y-%m-%d')
            
            # Create a pivot table for the timeline of demand
            pivot_df = pd.pivot_table(
                merged_data,
                values='Discrete Qty',
                index=['Plant', 'ps_par', 'ps_comp', 'pt_desc1', 'pt_desc2', 'ps_qty_per', 'po_vend', 'pt_prod_line', 'pt_dsgn_grp', 'pt_vend', 'pt_buyer', 'pod__chr08'],
                columns='Date_Formatted',
                aggfunc='sum',
                fill_value=0
            )
            
            # Reset index to convert back to a regular DataFrame
            pivot_df = pivot_df.reset_index()
            
            # Calculate component demand for each date
            logger.info("\nCalculating component demand for each date...")
            
            # Get the date columns
            date_columns = [col for col in pivot_df.columns if col not in 
                           ['Plant', 'ps_par', 'ps_comp', 'pt_desc1', 'pt_desc2', 'ps_qty_per', 'po_vend', 'pt_prod_line', 'pt_dsgn_grp', 'pt_vend', 'pt_buyer', 'pod__chr08']]
            
            # Calculate component demand by multiplying quantity by ps_qty_per
            for date_col in date_columns:
                pivot_df[f'Demand_{date_col}'] = pivot_df[date_col] * pivot_df['ps_qty_per']
            
            # Create a summary dataframe for component demand by date
            component_demand = pd.DataFrame()
            
            # Group by component and plant
            grouped = pivot_df.groupby(['Plant', 'ps_comp', 'pt_desc1', 'po_vend', 'pt_prod_line', 'pt_dsgn_grp', 'pt_vend', 'pt_buyer', 'pod__chr08'])
            
            # Create summary rows for each component
            component_rows = []
            
            for (plant, component, desc, vendor, prod_line, design_group, pt_vend, pt_buyer, pod_chr08), group in grouped:
                row = {
                    'Plant': plant,
                    'Component': component,
                    'Description': desc,
                    'Vendor': vendor,
                    'Product_Line': prod_line,
                    'Design_Group': design_group,
                    'PT_Vend': pt_vend,
                    'PT_Buyer': pt_buyer,
                    'POD_CHR08': pod_chr08
                }
                
                # Add demand for each date
                for date_col in date_columns:
                    demand_col = f'Demand_{date_col}'
                    row[date_col] = group[demand_col].sum()
                
                # Add total demand
                row['Total_Demand'] = sum(row[date_col] for date_col in date_columns)
                
                component_rows.append(row)
            
            # Create component demand dataframe
            component_demand = pd.DataFrame(component_rows)
            
            # Sort by total demand (descending)
            component_demand = component_demand.sort_values('Total_Demand', ascending=False)
            
            # Create inconsistency report
            logger.info("\nCreating inconsistency report...")
            
            # Filter rows where pt_vend <> po_vend or pt_buyer <> pod__chr08
            inconsistent_data = component_demand[
                (component_demand['PT_Vend'] != component_demand['Vendor']) | 
                (component_demand['PT_Buyer'] != component_demand['POD_CHR08'])
            ].copy()
            
            # Add inconsistency flags for clarity
            inconsistent_data['Vendor_Mismatch'] = inconsistent_data['PT_Vend'] != inconsistent_data['Vendor']
            inconsistent_data['Buyer_Mismatch'] = inconsistent_data['PT_Buyer'] != inconsistent_data['POD_CHR08']
            
            # Sort by component
            inconsistent_data = inconsistent_data.sort_values(['Component', 'Plant'])
            
            # Create summary dashboards
            logger.info("\nCreating summary dashboards...")
            
            # Summary by Vendor
            vendor_summary = component_demand.groupby('Vendor')['Total_Demand'].sum().reset_index()
            vendor_summary = vendor_summary.sort_values('Total_Demand', ascending=False)
            
            # Summary by Product Line
            product_line_summary = component_demand.groupby('Product_Line')['Total_Demand'].sum().reset_index()
            product_line_summary = product_line_summary.sort_values('Total_Demand', ascending=False)
            
            # Summary by Design Group
            design_group_summary = component_demand.groupby('Design_Group')['Total_Demand'].sum().reset_index()
            design_group_summary = design_group_summary.sort_values('Total_Demand', ascending=False)
            
            # Combined summary (Vendor + Product Line)
            combined_summary = component_demand.groupby(['Vendor', 'Product_Line'])['Total_Demand'].sum().reset_index()
            combined_summary = combined_summary.sort_values('Total_Demand', ascending=False)
            
            logger.info("\nComponent demand calculation completed")
            
            # Return all the dataframes for reporting
            return {
                'pivot_table': pivot_df,
                'component_demand': component_demand,
                'vendor_summary': vendor_summary,
                'product_line_summary': product_line_summary,
                'design_group_summary': design_group_summary,
                'combined_summary': combined_summary,
                'inconsistent_data': inconsistent_data
            }
        else:
            logger.error("Column 'Item Number' not found in demand data")
            return None
            
    except Exception as e:
        logger.error(f"Error analyzing demand with BOM: {str(e)}")
        logger.error(f"Stack trace:", exc_info=True)
        return None

def save_results(results, output_path):
    """
    Save analysis results to Excel file
    """
    try:
        logger.info(f"Saving results to: {output_path}")
        
        # Create output directory if it doesn't exist
        output_dir = os.path.dirname(output_path)
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
            logger.info(f"Created output directory: {output_dir}")
        
        # Create a writer object
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            # Write the component demand summary to the Excel file
            component_demand = results['component_demand']
            component_demand.to_excel(writer, sheet_name='Component Demand', index=False)
            
            # Write the pivot table to the Excel file
            pivot_df = results['pivot_table']
            pivot_df.to_excel(writer, sheet_name='Demand Timeline', index=False)
            
            # Write the summary dashboards
            results['vendor_summary'].to_excel(writer, sheet_name='Vendor Summary', index=False)
            results['product_line_summary'].to_excel(writer, sheet_name='Product Line Summary', index=False)
            results['design_group_summary'].to_excel(writer, sheet_name='Design Group Summary', index=False)
            results['combined_summary'].to_excel(writer, sheet_name='Combined Summary', index=False)
            
            # Write the inconsistency report
            if 'inconsistent_data' in results and not results['inconsistent_data'].empty:
                results['inconsistent_data'].to_excel(writer, sheet_name='Inconsistency Report', index=False)
                logger.info(f"Found {len(results['inconsistent_data'])} rows with inconsistencies")
            else:
                # Create an empty sheet with a message if no inconsistencies found
                pd.DataFrame({'Message': ['No inconsistencies found']}).to_excel(
                    writer, sheet_name='Inconsistency Report', index=False
                )
                logger.info("No inconsistencies found")
            
            # Auto-adjust column widths for component demand sheet
            for column in component_demand:
                column_width = max(component_demand[column].astype(str).map(len).max(), len(str(column))) + 2
                col_idx = component_demand.columns.get_loc(column)
                # Ensure we don't exceed Excel's column limit
                if col_idx < 26:
                    col_letter = chr(65 + col_idx)
                    writer.sheets['Component Demand'].column_dimensions[col_letter].width = column_width
            
            # Auto-adjust column widths for summary sheets
            for sheet_name in ['Vendor Summary', 'Product Line Summary', 'Design Group Summary', 'Combined Summary']:
                sheet_df = results[sheet_name.lower().replace(' ', '_')]
                for column in sheet_df:
                    column_width = max(sheet_df[column].astype(str).map(len).max(), len(str(column))) + 2
                    col_idx = sheet_df.columns.get_loc(column)
                    if col_idx < 26:
                        col_letter = chr(65 + col_idx)
                        writer.sheets[sheet_name].column_dimensions[col_letter].width = column_width
            
            # Auto-adjust column widths for inconsistency report
            if 'inconsistent_data' in results and not results['inconsistent_data'].empty:
                inconsistent_data = results['inconsistent_data']
                for column in inconsistent_data:
                    column_width = max(inconsistent_data[column].astype(str).map(len).max(), len(str(column))) + 2
                    col_idx = inconsistent_data.columns.get_loc(column)
                    if col_idx < 26:
                        col_letter = chr(65 + col_idx)
                        writer.sheets['Inconsistency Report'].column_dimensions[col_letter].width = column_width
        
        logger.info("Results saved successfully")
        return True
    except Exception as e:
        logger.error(f"Error saving results: {str(e)}")
        logger.error(f"Stack trace:", exc_info=True)
        return False

def main():
    try:
        # Parse command line arguments
        args = parse_arguments()
        
        # Set logging level based on verbose flag
        if args.verbose:
            logging.getLogger().setLevel(logging.DEBUG)
            logger.debug("Verbose logging enabled")
        
        # Step 1: Get the latest Excel file
        excel_file = get_latest_excel_file(args.excel_dir)
        if not excel_file:
            logger.error("Could not find Excel file. Exiting.")
            return
        
        # Step 2: Read customer demand data from Excel
        demand_df = read_excel_data(excel_file)
        if demand_df is None:
            logger.error("Failed to read demand data. Exiting.")
            return
        
        # Step 3: Execute SQL query to get BOM data
        bom_df = execute_sql_query(args.sql_file, args.db_server, args.db_name)
        if bom_df is None:
            logger.error("Failed to get BOM data. Exiting.")
            return
        
        # Step 4: Analyze demand with BOM data
        results = analyze_demand_with_bom(demand_df, bom_df, args.verbose)
        if results is None:
            logger.error("Failed to analyze demand with BOM. Exiting.")
            return
        
        # Step 5: Save results
        if save_results(results, args.output):
            logger.info(f"Analysis completed successfully. Results saved to {args.output}")
        else:
            logger.error("Failed to save results.")
    
    except Exception as e:
        logger.error(f"An error occurred: {str(e)}")
        logger.error(f"Stack trace:", exc_info=True)

if __name__ == "__main__":
    main()
