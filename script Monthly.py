import pandas as pd

# Load the database (pallet data) file
pallet_data = pd.read_excel('Pallet data.xlsx', usecols=[1, 9, 11], header=None, names=['SKU', 'PalletQty', 'PalletWeight'])

# Load the input Excel file
input_file = 'input.xlsx'
xls = pd.ExcelFile(input_file)

max_weight = 40000
max_pallets = 45

def process_sheet(sheet):
    df = pd.read_excel(xls, sheet_name=sheet, header=None)

    # Check if the second row contains the headers
    headers = df.iloc[1].values

    # Find the pattern 'SKU', 'Volume' (optional), 'Qty'
    pattern = ['SKU', 'Volume', 'Qty']
    pattern_without_volume = ['SKU', 'Qty']

    # Initialize variables to track the number of tables and their positions
    num_tables = 0
    table_positions = []

    # Iterate over the headers to find the pattern
    for i in range(len(headers)):
        if headers[i] == 'SKU':
            if i + 2 < len(headers) and headers[i + 1] == 'Volume' and headers[i + 2] == 'Qty':
                num_tables += 1
                table_positions.append(i)
            elif i + 1 < len(headers) and headers[i + 1] == 'Qty':
                num_tables += 1
                table_positions.append(i)

    # Determine the number of columns in each table
    num_columns = 3 if 'Volume' in headers else 2

    # Create a new dataframe for output
    output_df = pd.DataFrame()

    # Process each table
    for table_index, table_pos in enumerate(table_positions):
        # Extract the table data
        table_data = df.iloc[:, table_pos:table_pos + num_columns]

        # Process the current table
        current_order = {'weight': 0, 'pallets': 0, 'data': []}
        order_index = 1
        orders = []

        # Process each row in the table
        for i, row in table_data.iterrows():
            if row.iloc[0] == 'SKU':
                # Close and save the previous order if it has data
                if current_order['data']:
                    orders.append(current_order)
                    current_order = {'weight': 0, 'pallets': 0, 'data': []}
                    order_index = 1
                continue

            sku = row.iloc[0]
            qty = row.iloc[-1]

            # Find the pallet quantity and weight
            pallet_info = pallet_data[pallet_data['SKU'] == sku]
            if not pallet_info.empty:
                pallet_qty = pallet_info.iloc[0]['PalletQty']
                pallet_weight = pallet_info.iloc[0]['PalletWeight']

                # Calculate the number of pallets to ship
                ship_qty = (qty // pallet_qty) * pallet_qty
                if ship_qty == 0:
                    continue

                pallets = ship_qty // pallet_qty
                weight = pallets * pallet_weight

                # Check if the current order exceeds limits
                if (current_order['weight'] + weight > max_weight) or (current_order['pallets'] + pallets > max_pallets):
                    # Calculate how many pallets can fit in the current order
                    weight_space = max_weight - current_order['weight']
                    pallet_space = max_pallets - current_order['pallets']
                    
                    pallets_to_add = min(pallets, pallet_space, weight_space // pallet_weight)
                    
                    if pallets_to_add > 0:
                        # Add as many pallets as possible to the current order
                        current_order['weight'] += pallets_to_add * pallet_weight
                        current_order['pallets'] += pallets_to_add
                        current_order['data'].append((sku, pallets_to_add * pallet_qty, pallets_to_add, order_index, i))
                    
                    # Finalize the current order
                    orders.append(current_order)
                    
                    # Start a new order with the remaining pallets
                    remaining_pallets = pallets - pallets_to_add
                    order_index += 1
                    current_order = {
                        'weight': remaining_pallets * pallet_weight,
                        'pallets': remaining_pallets,
                        'data': [(sku, remaining_pallets * pallet_qty, remaining_pallets, order_index, i)] if remaining_pallets > 0 else []
                    }
                else:
                    # Update the current order
                    current_order['weight'] += weight
                    current_order['pallets'] += pallets
                    current_order['data'].append((sku, ship_qty, pallets, order_index, i))

        # Append the last order if it has data
        if current_order['data']:
            orders.append(current_order)

        # Create output columns for the current table
        max_orders = max(order['data'][-1][3] for order in orders) if orders else 0
        output_columns = pd.DataFrame(index=table_data.index)
        
        for i in range(1, max_orders + 1):
            output_columns[f'Order {i} Qty'] = ''
            output_columns[f'Order {i} Pallets'] = ''

        # Fill the output columns
        for order in orders:
            for item in order['data']:
                sku, ship_qty, pallets, index, row_num = item
                output_columns.loc[row_num, f'Order {index} Qty'] = ship_qty
                output_columns.loc[row_num, f'Order {index} Pallets'] = pallets

        # Add order headers to the 'SKU' row
        sku_row_index = table_data.index[table_data.iloc[:, 0] == 'SKU'][0]
        for i in range(1, max_orders + 1):
            output_columns.loc[sku_row_index, f'Order {i} Qty'] = f'Order {i}'
            output_columns.loc[sku_row_index, f'Order {i} Pallets'] = 'Pallets'

        # Combine the original table with its output
        combined_table = pd.concat([table_data, output_columns], axis=1)
        
        # Add the combined table to the output dataframe
        output_df = pd.concat([output_df, combined_table], axis=1)

    return output_df

# Process each sheet in the Excel file
output_sheets = {}
for sheet in xls.sheet_names:
    output_sheets[sheet] = process_sheet(sheet)

# Save the result to a new Excel file
with pd.ExcelWriter('output.xlsx', engine='openpyxl') as writer:
    for sheet_name, df in output_sheets.items():
        df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
