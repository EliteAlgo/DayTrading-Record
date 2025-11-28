import pandas as pd
import os

def process_portfolio():
    input_file = 'S1-11-27-SUMMARY.xlsx'
    sheet_name = 'Portfolios'
    output_file = 'portfolio_summary.xlsx'

    if not os.path.exists(input_file):
        print(f"Error: File {input_file} not found.")
        return

    try:
        df = pd.read_excel(input_file, sheet_name=sheet_name)
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return

    # Ensure required columns exist
    required_columns = ['User ID', 'Portfolio Name', 'PNL', 'Strategy Tag']
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        print(f"Error: Missing columns: {missing_columns}")
        print(f"Available columns: {df.columns.tolist()}")
        return

    # Create Portfolio Group
    df['Portfolio Group'] = df['Portfolio Name'].astype(str).str[:5]

    # Create Pivot Table
    # Rows: Portfolio Group, Strategy Tag
    # Columns: User ID
    # Values: PNL
    pivot_df = df.pivot_table(index=['Portfolio Group', 'Strategy Tag'], 
                              columns='User ID', 
                              values='PNL', 
                              aggfunc='sum')
    
    # Reset index to make Portfolio Group and Strategy Tag regular columns
    summary = pivot_df.reset_index()

    # Fill NaNs with empty string (or keep as NaN for Excel, but user asked for blank)
    # For calculation/display purposes, NaN is often better, but for "blank" visual:
    summary = summary.fillna('')

    # Format for display
    pd.set_option('display.max_rows', None)
    pd.set_option('display.width', 1000)
    print("Summary Data:")
    print(summary)

    # Save to Excel
    try:
        summary.to_excel(output_file, index=False)
        print(f"\nSummary saved to {output_file}")
    except Exception as e:
        print(f"Error saving to Excel: {e}")

    # Save as PNG
    try:
        import matplotlib.pyplot as plt
        from pandas.plotting import table

        # Calculate figure size
        num_rows = len(summary)
        num_cols = len(summary.columns)
        
        # Dynamic sizing
        width = max(10, num_cols * 1.5)
        height = max(4, num_rows * 0.4 + 1)

        fig, ax = plt.subplots(figsize=(width, height))
        ax.axis('off')
        
        # Create table
        tbl = table(ax, summary, loc='center', cellLoc='center')
        
        # Style the table
        tbl.auto_set_font_size(False)
        tbl.set_fontsize(10)
        tbl.scale(1.2, 1.2)
        
        # Adjust column widths if needed (optional)
        # for key, cell in tbl.get_celld().items():
        #     cell.set_linewidth(0.5)

        png_file = 'portfolio_summary.png'
        plt.savefig(png_file, bbox_inches='tight', pad_inches=0.1)
        print(f"Summary image saved to {png_file}")
    except ImportError:
        print("matplotlib not installed, skipping image generation.")
    except Exception as e:
        print(f"Error saving image: {e}")

if __name__ == "__main__":
    process_portfolio()
