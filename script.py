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

    # List to hold all order data
    orders = []

    # Variables to track the current order
    current_order = {'weight': 0, 'pallets': 0, 'data': []}
    order_index = 1
    
    qty_column_index = None  # To determine where the QTY column is

    for i, row in df.iterrows():
        if 'Qty' in row.values:
            # Determine the column index of 'QTY'
            qty_column_index = row[row == 'Qty'].index[0]

            # If a new QTY header is found, reset the order index and process the previous order if any
            if current_order['data']:
                orders.append(current_order)
            current_order = {'weight': 0, 'pallets': 0, 'data': []}
            order_index = 1
            continue
        
        # If the row contains SKU data (excluding headers), process it
        if isinstance(row[0], (int, float)):
            sku = row[0]
            qty = row[qty_column_index]

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
                if current_order['weight'] + weight > max_weight or current_order['pallets'] + pallets > max_pallets:
                    # Finalize the current order and start a new one
                    orders.append(current_order)
                    current_order = {'weight': 0, 'pallets': 0, 'data': []}
                    order_index += 1

                # Update the current order
                current_order['weight'] += weight
                current_order['pallets'] += pallets
                current_order['data'].append((sku, ship_qty, pallets, order_index))

    # Append the last order if it has data
    if current_order['data']:
        orders.append(current_order)

    # Create a new dataframe for output
    output_df = df.copy()

    for order in orders:
        for item in order['data']:
            sku, ship_qty, pallets, index = item
            output_df.loc[output_df[0] == sku, f'Order {index}'] = ship_qty

    return output_df

# Process each sheet in the Excel file
output_sheets = {}
for sheet in xls.sheet_names:
    output_sheets[sheet] = process_sheet(sheet)

# Save the result to a new Excel file
with pd.ExcelWriter('output.xlsx', engine='openpyxl') as writer:
    for sheet_name, df in output_sheets.items():
        df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
