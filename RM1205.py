import streamlit as st
import pandas as pd
from io import BytesIO
import xlsxwriter
from xlsxwriter.utility import xl_col_to_name
import datetime

st.set_page_config(page_title="Rate Maker")

# Fixed costs and other constants
WORKING_DAYS_PER_YEAR = 252
MANAGEMENT_COST_PERCENTAGE = 0.0650
FACILITIES_COST_PERCENTAGE = 0.0645
ADMIN_COST_PERCENTAGE = 0.1475

# Default margins for specified brackets, adjustable by the user

if 'custom_margins' not in st.session_state:
    st.session_state['custom_margins'] = {
        (1, 10): 5,
        (11, 25): 7,
        (26, 75): 10,
        (76, 250): 15
    }
# Constants for daily rates based on weight categories
WEIGHT_RATES = {
    "End to End": {
        "0 - 10": 5,
        "11 - 24": 10,
        ">=25": 15
    },
    "End to End without Pickup": {
        "0 - 10": 4.5,
        "11 - 24": 9,
        ">=25": 17.5
    },
    "Final Mile Only": {
        "0 - 10": 4,
        "11 - 24": 8,
        ">=25": 19
    }
}
# Terminals and their zone details
terminals = {
    '50': 'SLOK(Backyard)',
    '100': 'SLOK',
    '110': 'SLHA/PK/BA',
    '120': 'SLLN/WS/OR/HV',
    '130': 'SLMT',
    '135': 'SLOT',
    '140': 'SLQC'
}

# Inverting the dictionary to map from zone names to zone codes
name_to_zone_number = {v: k for k, v in terminals.items()}

# Creating a sorted list of terminal names for the dropdown to display terminal names
terminal_names = sorted(name_to_zone_number.keys())

# First mile zone details
firstmile_zone_details = {
    '50': {
        'pickup_cost': 350.0000,
        'gaylords_per_truck': 18.0000,
        'pcs_gaylord': {
            '4lbs': 125.0000,
            '13lbs':65.0000,
            '>=75': 8.0000
        }
    },
    '100': {
        'pickup_cost': 450.0000,
        'gaylords_per_truck': 18.0000,
        'pcs_gaylord': {
            '4lbs': 125.0000,
            '13lbs': 65.0000,
            '>=75': 8.0000
        }
    },
    '110': {
        'pickup_cost': 650.0000,
        'gaylords_per_truck': 18.0000,
        'pcs_gaylord': {
            '4lbs': 125.0000,
            '13lbs': 65.0000,
            '>=75': 08.0000
        }
    },
    '120': {
        'pickup_cost': 900.0000,
        'gaylords_per_truck': 18.0000,
        'pcs_gaylord': {
            '4lbs': 125.0000,
            '13lbs': 65.0000,
            '>=75': 08.0000
        }
    },
    '130': {
        'pickup_cost': 1300.0000,
        'gaylords_per_truck': 18.0000,
        'pcs_gaylord': {
            '4lbs': 125.0000,
            '13lbs': 65.0000,
            '>=75': 08.0000
        }
    },
    '135': {
        'pickup_cost': 1300.0000,
        'gaylords_per_truck': 18.0000,
        'pcs_gaylord': {
            '4lbs': 125.0000,
            '13lbs': 65.0000,
            '>=75': 08.0000
        }
    },
    '140': {
        'pickup_cost': 1300.0000,
        'gaylords_per_truck': 18.0000,
        'pcs_gaylord': {
            '4lbs': 125.0000,
            '13lbs': 65.0000,
            '>=75': 08.0000
        }
    }
}

# Middle mile Zone details 
middle_mile_pickup_details = {
    '50': {
        'KM Radius': 200,
        'Pickup Cost': 0.0000,
        '# of Gaylords / Truck': 20.0000,
        'PCs/Gaylord - 4LBs': 150.0000,
        'PCs/Gaylord - 13 LBs': 75.0000,
        'PCs/Gaylord > = 75 Lbs': 10.0000
    },
    '100': {
        'KM Radius': 200,
        'Pickup Cost': 0.0000,
        '# of Gaylords / Truck': 20.0000,
        'PCs/Gaylord - 4LBs': 150.0000,
        'PCs/Gaylord - 13 LBs': 75.0000,
        'PCs/Gaylord > = 75 Lbs': 10.0000
    },
    '110': {
        'KM Radius': 350,
        'Pickup Cost': 450.0000,
        '# of Gaylords / Truck': 20.0000,
        'PCs/Gaylord - 4LBs': 150.0000,
        'PCs/Gaylord - 13 LBs': 75.0000,
        'PCs/Gaylord > = 75 Lbs': 10.0000
    },
    '120': {
        'KM Radius': 500,
        'Pickup Cost': 800.0000,
        '# of Gaylords / Truck': 20.0000,
        'PCs/Gaylord - 4LBs': 150.0000,
        'PCs/Gaylord - 13 LBs': 75.0000,
        'PCs/Gaylord > = 75 Lbs': 10.0000
    },
    '130': {
        'KM Radius': 650,
        'Pickup Cost': 1200.0000,
        '# of Gaylords / Truck': 20.0000,
        'PCs/Gaylord - 4LBs': 150.0000,
        'PCs/Gaylord - 13 LBs': 75.0000,
        'PCs/Gaylord > = 75 Lbs': 10.0000
    },
    '135': {
        'KM Radius': 650,
        'Pickup Cost': 1400.0000,
        '# of Gaylords / Truck': 18.0000,
        'PCs/Gaylord - 4LBs': 150.0000,
        'PCs/Gaylord - 13 LBs': 75.0000,
        'PCs/Gaylord > = 75 Lbs': 10.0000
    },
    '140': {
        'KM Radius': 650,
        'Pickup Cost': 1700.0000,
        '# of Gaylords / Truck': 20.0000,
        'PCs/Gaylord - 4LBs': 150.0000,
        'PCs/Gaylord - 13 LBs': 75.0000,
        'PCs/Gaylord > = 75 Lbs': 10.0000
    }
}

# First Mile Variance Details 
firstmile_variance_details = { '50': {
        'upto_13_lbs': 0.0205,
        'upto_75_lbs': 0.0442
    },
    '100': {
        'upto_13_lbs': 0.0205,
        'upto_75_lbs': 0.0442
    },
    '110': {
        'upto_13_lbs': 0.0296,
        'upto_75_lbs': 0.0638
    },
    '120': {
        'upto_13_lbs': 0.0410,
        'upto_75_lbs': 0.0884
    },
    '130': {
        'upto_13_lbs': 0.0593,
        'upto_75_lbs': 0.1277
    }, 
    '135': {
        'upto_13_lbs': 0.0593,
        'upto_75_lbs': 0.1277
    }, 
    '140': {
        'upto_13_lbs': 0.0593,
        'upto_75_lbs': 0.1277
    } } 


#   Middle mile variance details
middle_mile_variance_details = {
    'SLOK(Backyard)': {
        'Upto 13 lbs': 0.0000,
        'Upto 75 lbs': 0.0000
    },
    'SLOK': {
        'Upto 13 lbs': 0.0000,
        'Upto 75 lbs': 0.0000
    },
    'SLHA/PK/BA': {
        'Upto 13 lbs': 0.0167,
        'Upto 75 lbs': 0.0315
    },
    'SLLN/WS/OR/HV': {
        'Upto 13 lbs': 0.0296,
        'Upto 75 lbs': 0.0559
    },
    'SLMT': {
        'Upto 13 lbs': 0.0444,
        'Upto 75 lbs': 0.0839
    },
    'SLOT': {
        'Upto 13 lbs': 0.0576,
        'Upto 75 lbs': 0.1087
    },
    'SLQC': {
        'Upto 13 lbs': 0.0630,
        'Upto 75 lbs': 0.1188
    }
}

# Final Mile Variance Details 
final_mile_variance_rates = {
    'R1': {
        'SLOK(Backyard)': {'Upto 13 lbs': 0.0111, 'Upto 75 lbs': 0.0806},
        'SLOK': {'Upto 13 lbs': 0.0167, 'Upto 75 lbs': 0.1089},
        'SLHA/PK/BA': {'Upto 13 lbs': 0.0167, 'Upto 75 lbs': 0.1089},
        'SLLN/WS/OR/HV': {'Upto 13 lbs': 0.0167, 'Upto 75 lbs': 0.1089},
        'SLMT': {'Upto 13 lbs': 0.0167, 'Upto 75 lbs': 0.1065},
        'SLOT': {'Upto 13 lbs': 0.0167, 'Upto 75 lbs': 0.1065},
        'SLQC': {'Upto 13 lbs': 0.0167, 'Upto 75 lbs': 0.1065},
    },
    'R2': {
        'SLOK': {'Upto 13 lbs': 0.05556, 'Upto 75 lbs': 0.13710},
        'SLHA/PK/BA': {'Upto 13 lbs': 0.05556, 'Upto 75 lbs': 0.13710},
        'SLLN/WS/OR/HV': {'Upto 13 lbs': 0.05556, 'Upto 75 lbs': 0.13710},
        'SLMT': {'Upto 13 lbs': 0.05556, 'Upto 75 lbs': 0.13710},
        'SLOT': {'Upto 13 lbs': 0.05556, 'Upto 75 lbs': 0.13710},
        'SLQC': {'Upto 13 lbs': 0.05556, 'Upto 75 lbs': 0.13710},
    }
}

# First mile sort costs and Final mile sort costs
First_sort_costs = {'SLOK(Backyard)': {'1-4': 0.3000, '5-13': 0.4000, '14-250': 1.5000},
    'SLOK': {'1-4': 0.300, '5-13': 0.4000, '14-250': 1.5000},
    'SLHA/PK/BA': {'1-4': 0.3000, '5-13': 0.4000, '14-250': 1.5000},
    'SLLN/WS/OR/HV': {'1-4': 0.3000, '5-13': 0.4000, '14-250': 1.5000},
    'SLMT': {'1-4': 0.3000, '5-13': 0.4000, '14-250': 1.5000},
    'SLOT': {'1-4': 0.3000, '5-13': 0.4000, '14-250': 1.5000},
    'SLQC': {'1-4': 0.3000, '5-13': 0.4000, '14-250': 1.5000},}  

Final_sort_costs = {
    'SLOK(Backyard)': {'1-4': 0.2000, '5-13': 0.2500, '14-250': 1.5000},
    'SLOK': {'1-4': 0.2000, '5-13': 0.2500, '14-250': 1.5000},
    'SLHA/PK/BA': {'1-4': 0.2000, '5-13': 0.2500, '14-250': 1.5000},
    'SLLN/WS/OR/HV': {'1-4': 0.2000, '5-13': 0.2500, '14-250': 1.5000},
    'SLMT': {'1-4': 0.2000, '5-13': 0.2500, '14-250':1.5000},
    'SLOT': {'1-4': 0.2000, '5-13': 0.2500, '14-250':1.5000},
    'SLQC': {'1-4': 0.2000, '5-13': 0.2500, '14-250':1.5000},
}
final_mile_costs = {
    'R1': {
        'SLOK(Backyard)': {'1-4': 1.9000, '5-13': 2.0000, '14-250': 7.0000},
        'SLOK': {'1-4': 2.1000, '5-13': 2.2500, '14-250': 9.0000},
        'SLHA/PK/BA': {'1-4': 2.1000, '5-13': 2.2500, '14-250': 9.0000},
        'SLLN/WS/OR/HV': {'1-4': 2.1000, '5-13': 2.2500, '14-250': 9.0000},
        'SLMT': {'1-4': 2.2500, '5-13': 2.4000, '14-250': 9.0000},
        'SLOT': {'1-4': 2.2500, '5-13': 2.4000, '14-250': 9.0000},
        'SLQC': {'1-4': 2.2500, '5-13': 2.4000, '14-250': 9.0000},
    },
    'R2': {
        'SLOK': {'1-4': 3.0000, '5-13': 3.5000, '14-250': 12.0000},
        'SLHA/PK/BA': {'1-4': 3.0000, '5-13': 3.5000, '14-250': 12.0000},
        'SLLN/WS/OR/HV': {'1-4': 3.0000, '5-13': 3.5000, '14-250': 12.0000},
        'SLMT': {'1-4': 3.0000, '5-13': 3.5000, '14-250': 12.0000},
        'SLOT': {'1-4': 3.0000, '5-13': 3.5000, '14-250': 12.0000},
        'SLQC': {'1-4': 3.0000, '5-13': 3.5000, '14-250': 12.0000},
    }
}

# Cost Functions # First Mile pickup Cost 
def calculate_pickup_cost_with_variance(weight, user_selected_pickup_details, zone_variance_details):
    pcs_gaylord = user_selected_pickup_details['pcs_gaylord']
    gaylords_per_truck = user_selected_pickup_details['gaylords_per_truck']
    pickup_cost = user_selected_pickup_details['pickup_cost']
    # Calculate base costs for specific weight points
    base_cost_for_4 = pickup_cost / (gaylords_per_truck * pcs_gaylord['4lbs'])
    base_cost_for_13 = pickup_cost / (gaylords_per_truck * pcs_gaylord['13lbs'])
    if weight == 4:
        base_cost = base_cost_for_4
    elif weight == 13:
        base_cost = base_cost_for_13
    elif weight >= 75:
        base_cost = pickup_cost / (gaylords_per_truck * pcs_gaylord['>=75'])
    elif weight < 4:
        variance = zone_variance_details['upto_13_lbs']
        base_cost = base_cost_for_4 - ((4 - weight) * variance)
    elif 14 <= weight < 75:
        variance = zone_variance_details['upto_75_lbs']
        base_cost = base_cost_for_13 + ((weight - 13) * variance)
    else:  # For weights between 5 and 12
        variance = zone_variance_details['upto_13_lbs']
        base_cost = base_cost_for_4 + ((weight - 4) * variance)
    return round(base_cost, 6)

# Calculating sort cost
def calculate_first_sort_cost(terminal, weight):
    if weight <= 4:
        return First_sort_costs[terminal]['1-4']
    elif weight <= 13:
        return First_sort_costs[terminal]['5-13']
    else:  
        return First_sort_costs[terminal]['14-250']
    
    # Middle mile cost calculation  
def calculate_middle_mile_cost_with_variance(zone_code, weight, middle_mile_pickup_details, middle_mile_variance_details):
    terminal_name = terminals[zone_code]
    mm_pickup_cost = middle_mile_pickup_details[zone_code]['Pickup Cost']
    gaylords_per_truck = middle_mile_pickup_details[zone_code]['# of Gaylords / Truck']
    # Getting the pieces per gaylord values for different weights - mainly for 4 lbs, 13lbs and 75 lbs.
    pcs_gaylord_4lbs = middle_mile_pickup_details[zone_code]['PCs/Gaylord - 4LBs']
    pcs_gaylord_13lbs = middle_mile_pickup_details[zone_code]['PCs/Gaylord - 13 LBs']
    pcs_gaylord_75lbs = middle_mile_pickup_details[zone_code]['PCs/Gaylord > = 75 Lbs']
    # Calculates the base cost for 4 lbs and 13 lbs weights
    base_cost_for_4 = mm_pickup_cost / (gaylords_per_truck * pcs_gaylord_4lbs)
    base_cost_for_13 = mm_pickup_cost / (gaylords_per_truck * pcs_gaylord_13lbs)
    variance_upto_13_lbs = middle_mile_variance_details[terminal_name].get('Upto 13 lbs', 0)
    variance_upto_75_lbs = middle_mile_variance_details[terminal_name].get('Upto 75 lbs', 0)
    # Initializing  base_cost to ensure it has a value 
    base_cost = 0
    if weight == 4:
        base_cost = base_cost_for_4
    elif weight < 4:
        base_cost = (base_cost_for_4 - ((4 - weight) * variance_upto_13_lbs))
    elif 5 <= weight < 13:
    # For weights between 5 and 12, it starts with the base cost for 4 and add variance up to 13 lbs
        base_cost = base_cost_for_4 + ((weight - 4) * variance_upto_13_lbs)
    elif weight == 13:
        base_cost = base_cost_for_13
    elif 14 <= weight < 75:
        base_cost = base_cost_for_13 + ((weight - 13) * variance_upto_75_lbs)
    elif weight >= 75:
        base_cost = mm_pickup_cost / (gaylords_per_truck * pcs_gaylord_75lbs)
    
    return round(base_cost, 6)

# Final sort cost calculation 
def calculate_final_sort_cost(terminal, weight):
    if weight <= 4:
        return Final_sort_costs[terminal]['1-4']
    elif weight <= 13:
        return Final_sort_costs[terminal]['5-13']
    else:  # Here weight is always less than or equal to 250
        return Final_sort_costs[terminal]['14-250']
    
# Final mile cost calculation    
def calculate_final_mile_cost_with_variance(terminal, rate_type, weight, final_mile_costs, final_mile_variance_rates):
    # Check if the terminal exists in the provided rate type; if not, return 0.
    if terminal not in final_mile_costs[rate_type]:
        print(f"Terminal '{terminal}' not found in rate type '{rate_type}'. Returning cost of 0.")
        return 0
    # Retrieve the cost brackets and variance details for the terminal and rate type.
    cost_brackets = final_mile_costs[rate_type][terminal]
    variance = final_mile_variance_rates[rate_type][terminal]
    # Calculate cost based on weight categories with variance adjustments.
    if weight <= 4:
        # For weights 1-4, use cost of 4 lbs minus variance based on deviation from 4 lbs.
        adjusted_cost = cost_brackets['1-4'] - (4 - weight) * variance['Upto 13 lbs']
    elif weight <= 13:
        # For weights 5-13, use the cost for 4 lbs as a base and add variance for each additional pound.
        adjusted_cost = cost_brackets['1-4'] + (weight - 4) * variance['Upto 13 lbs']
    elif 14 <= weight < 75:
        # Starting with the cost for 13 lbs, add incremental variance for each pound above 13 lbs up to 74 lbs.
        base_cost = cost_brackets['1-4'] + 9 * variance['Upto 13 lbs']  # This calculates the cost at 13 lbs
        variance_increment = variance.get('Upto 75 lbs', 0)
        adjusted_cost = base_cost + (weight - 13) * variance_increment
    else:
        # For weights above 75 lbs, use the maximum bracket cost.
        adjusted_cost = cost_brackets['14-250']
    # To ensure that adjusted_cost is never negative; it must always be at least 0.
    final_cost = max(0, adjusted_cost)
    return round(final_cost, 6)

# Total cost calculation Function
def calculate_costs(service_type, user_selected_zone, user_selected_pickup_details, firstmile_variance_details, firstmile_zone_details, final_mile_variance_rates, middle_mile_pickup_details, middle_mile_variance_details, final_mile_costs, custom_margins):
    selected_zone_details = firstmile_zone_details.get(user_selected_zone)
    if selected_zone_details is None:
        st.error(f"Selected pickup zone '{user_selected_zone}' is not recognized.")
        return None, None
    
    # Retrieve custom margins from the provided dictionary
    custom_margins_dict = {bracket: margin / 100 for bracket, margin in custom_margins.items()}
    
    costs_df = pd.DataFrame(index=pd.RangeIndex(start=1, stop=251, name='Weight in lbs'))
    
    for weight in range(1, 251):
        custom_margin = next((margin for (start, end), margin in custom_margins_dict.items() if start <= weight <= end), 0)
        sale_rate_factor = 1 - custom_margin
        margin_factor = 1 + custom_margin
        
        for zone_code, terminal_name in terminals.items():
            zone_variance_details = firstmile_variance_details[zone_code]
            pickup_cost = 0
            first_sort_cost = 0
            middle_mile_cost = 0
            final_sort_cost = 0
            final_mile_cost_r1 = 0
            final_mile_cost_r2 = 0
            total_direct_cost_r1 = 0
            total_direct_cost_r2 = 0
            
            if service_type == 'End to End':
                pickup_cost = calculate_pickup_cost_with_variance(weight, user_selected_pickup_details, firstmile_variance_details[user_selected_zone])
                first_sort_cost = calculate_first_sort_cost(terminal_name, weight)
                middle_mile_cost = calculate_middle_mile_cost_with_variance(zone_code, weight, middle_mile_pickup_details, middle_mile_variance_details)
                final_sort_cost = calculate_final_sort_cost(terminal_name, weight)
                final_mile_cost_r1 = calculate_final_mile_cost_with_variance(terminal_name, 'R1', weight, final_mile_costs, final_mile_variance_rates)
                if zone_code != '50':  # Zone 50 applies to R1 but not to R2
                    final_mile_cost_r2 = calculate_final_mile_cost_with_variance(terminal_name, 'R2', weight, final_mile_costs, final_mile_variance_rates)

            elif service_type == 'End to End without Pickup':
                first_sort_cost = calculate_first_sort_cost(terminal_name, weight)
                middle_mile_cost = calculate_middle_mile_cost_with_variance(zone_code, weight, middle_mile_pickup_details, middle_mile_variance_details)
                final_sort_cost = calculate_final_sort_cost(terminal_name, weight)
                final_mile_cost_r1 = calculate_final_mile_cost_with_variance(terminal_name, 'R1', weight, final_mile_costs, final_mile_variance_rates)
                if zone_code != '50':
                    final_mile_cost_r2 = calculate_final_mile_cost_with_variance(terminal_name, 'R2', weight, final_mile_costs, final_mile_variance_rates)

            elif service_type == 'Final Mile Only':
                final_sort_cost = calculate_final_sort_cost(terminal_name, weight)
                final_mile_cost_r1 = calculate_final_mile_cost_with_variance(terminal_name, 'R1', weight, final_mile_costs, final_mile_variance_rates)
                if zone_code != '50':
                    final_mile_cost_r2 = calculate_final_mile_cost_with_variance(terminal_name, 'R2', weight, final_mile_costs, final_mile_variance_rates)

            total_direct_cost_r1 = pickup_cost + first_sort_cost + middle_mile_cost + final_sort_cost + final_mile_cost_r1
            if final_mile_cost_r2:
                total_direct_cost_r2 = pickup_cost + first_sort_cost + middle_mile_cost + final_sort_cost + final_mile_cost_r2
                overhead_cost_r2 = total_direct_cost_r2 + (total_direct_cost_r2 / sale_rate_factor * (MANAGEMENT_COST_PERCENTAGE + FACILITIES_COST_PERCENTAGE + ADMIN_COST_PERCENTAGE))
                final_cost_r2 = overhead_cost_r2 * margin_factor
                costs_df.at[weight, f'{terminal_name} (Zone {zone_code}) - R2'] = final_cost_r2

            overhead_cost_r1 = total_direct_cost_r1 + (total_direct_cost_r1 / sale_rate_factor * (MANAGEMENT_COST_PERCENTAGE + FACILITIES_COST_PERCENTAGE + ADMIN_COST_PERCENTAGE))
            final_cost_r1 = overhead_cost_r1 * margin_factor
            costs_df.at[weight, f'{terminal_name} (Zone {zone_code}) - R1'] = final_cost_r1

    return costs_df


def to_excel(user_inputs_df,df1, df2):
    """Convert two dataframes into an Excel file, return the file content ready for download."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Define formats
        money_format = writer.book.add_format({'num_format': '$#,##0.00', 'align': 'right'})
        header_format = writer.book.add_format({'bold': True, 'bg_color': '#FFFF00', 'align': 'center'})
        index_format = writer.book.add_format({'align': 'left'})  # Format for 'Weight in lbs' column
        # Write user inputs to the first sheet
        user_inputs_df.to_excel(writer, sheet_name='User Inputs', index=False)

        # Configure each DataFrame into a separate sheet
        for name, df in [('R1', df1), ('R2', df2)]:
            worksheet = writer.book.add_worksheet(name)
            writer.sheets[name] = worksheet
            # Write headers
            worksheet.write(0, 0, 'Weight in lbs', header_format)
            for col_num, value in enumerate(df.columns, start=1):
                worksheet.write(0, col_num, value, header_format)
            # Write data with formatting
            for row_idx, row in enumerate(df.itertuples(), start=1):
                # Write index (Weight in lbs)
                worksheet.write(row_idx, 0, getattr(row, 'Index'), index_format)
                # Write other data with currency format
                for col_idx, value in enumerate(row[1:], start=1):
                    worksheet.write(row_idx, col_idx, value, money_format)
            # Set column widths more appropriately
            worksheet.set_column(0, 0, 15, index_format)  # Width for 'Weight in lbs' column
            worksheet.set_column(1, len(df.columns), 18, money_format)  # Adjust width for monetary columns

            # Adjust column names of df2 from starting with "1" to starting with "2"
            if name == 'R2':
                adjusted_columns = {
                    'SLOK (Zone 100) - R2': 'SLOK (Zone 200) - R2',
                    'SLHA/PK/BA (Zone 110) - R2': 'SLHA/PK/BA (Zone 210) - R2',
                    'SLLN/WS/OR/HV (Zone 120) - R2': 'SLLN/WS/OR/HV (Zone 220) - R2',
                    'SLMT (Zone 130) - R2': 'SLMT (Zone 230) - R2',
                    'SLOT (Zone 135) - R2': 'SLOT (Zone 235) - R2',
                    'SLQC (Zone 140) - R2': 'SLQC (Zone 240) - R2'
                }
                df2_adjusted = df2.rename(columns=adjusted_columns)
                for col_num, value in enumerate(df2_adjusted.columns, start=1):
                    worksheet.write(0, col_num, value, header_format)
                for row_idx, row in enumerate(df2_adjusted.itertuples(), start=1):
                    # Write index (Weight in lbs)
                    worksheet.write(row_idx, 0, getattr(row, 'Index'), index_format)
                    # Write other data with currency format
                    for col_idx, value in enumerate(row[1:], start=1):
                        worksheet.write(row_idx, col_idx, value, money_format)
                # Set column widths more appropriately
                worksheet.set_column(1, len(df2_adjusted.columns), 18, money_format)

    output.seek(0)
    return output.getvalue()


# Streamlit Application Layout
st.image('logo.png', width=600)
st.markdown(
    "<h1 style='text-align: center; color: grey;'>Final Mile Rate Maker</h1>",
    unsafe_allow_html=True
)
# service_type
service_type = None
zone_numbers = list(terminals.keys())

# Assuming st.session_state['custom_margins'] is initialized as shown in your code
weight_brackets = [(1, 10), (11, 25), (26, 75), (76, 250)]
bracket_descriptions = ["1-10 lbs", "11-25 lbs", "26-75 lbs", "76-250 lbs"]


# Streamlit form#1 for user input
with st.form("input_form"):
    st.write("Enter Opportunity Details")
    opportunity_name = st.text_input("Opportunity Name", key='opportunity_name')
    quote_prepared_by = st.text_input("Quote Prepared By", key='quote_prepared_by')
    freight_pickup_service = st.selectbox("Does the freight need to be picked up from the customer location to the induction site?", ['Yes', 'No'], key='freight_pickup_service')
    sort_initial_freight = st.selectbox("Is sorting of initial mixed freight at the induction site required?", ['Yes', 'No'], key='sort_initial_freight')
    pick_up_location = st.selectbox("Pickup Zone", options=terminal_names, key='pick_up_location')
    pick_up_schedule = st.text_input("Pickup Schedule", key='pick_up_schedule')
    pick_up_timings = st.text_input("Pickup Timings", key='pick_up_timings')
    average_shipment_weight = st.selectbox("Average Shipment Weight (In Pounds)", ["0 - 10", "11 - 24", ">=25"], key='average_shipment_weight')
    use_rate_shopping_system = st.radio("Do you use a rate shopping system?", ["Yes", "No"], key='rate_shopping_system')
    shipping_sla = st.selectbox("Choose Shipping SLA", ["Next Day", "Same Day", "2-days"], key='shipping_sla')
    service_type_required = st.radio("Choose Service Type Required", ["Non-dedicated", "Dedicated"], key='service_type_required')
    avg_shipments_per_day = st.number_input("Average Shipments Per Day", min_value=1, key='avg_shipments_per_day')
    avg_pieces_per_shipment = st.number_input("Average Pieces Per Shipment", min_value=1, key='avg_pieces_per_shipment')
    submit = st.form_submit_button("Submit")

    if submit:
        if not all([opportunity_name, quote_prepared_by, pick_up_schedule, pick_up_timings, avg_shipments_per_day, avg_pieces_per_shipment]):
            st.error("All fields are required.")
            st.stop()
        if shipping_sla == 'Same Day' and service_type_required == 'Dedicated':
            st.warning("Please see Admin for Same Day and Dedicated Service pricing")
            st.stop()
        elif shipping_sla == 'Same Day':
            st.warning("Please see Admin for Same Day Pricing Quotes")
            st.stop()
        elif service_type_required == 'Dedicated':
            st.warning("Please see Admin for Dedicated Service pricing")
            st.stop()
        user_selected_zone_name = pick_up_location
        user_selected_zone_number = name_to_zone_number[user_selected_zone_name]
        user_selected_pickup_details = firstmile_zone_details.get(user_selected_zone_number)
        if user_selected_pickup_details:
            service_type = 'End to End' if freight_pickup_service == 'Yes' and sort_initial_freight == 'Yes' else \
                           'End to End without Pickup' if sort_initial_freight == 'Yes' else 'Final Mile Only'
            st.write(f"Service Type: {service_type}")
            rate_per_piece = WEIGHT_RATES[service_type][average_shipment_weight]
            estimated_revenue = avg_shipments_per_day * avg_pieces_per_shipment * rate_per_piece * WORKING_DAYS_PER_YEAR
            st.metric("Estimated Annual Revenue", f"${estimated_revenue:,.0f}")
            st.write("Base freight revenue only. Fuel is billed separately.")
            st.write(f"Selected Sell Rate per Piece: ${rate_per_piece} based on average weight category '{average_shipment_weight}'")
            st.session_state['service_type'] = service_type
            st.session_state['user_selected_zone_number'] = user_selected_zone_number
            #st.session_state['target_margin'] = 0  # Initialize with default margin
        else:
            st.error("No pickup details available for the selected zone.")

# Custom Margin Form
with st.form("custom_margin_form"):
    st.write("Set Custom Margins for Weight Brackets (%):")
    for bracket, default_margin in st.session_state.get('custom_margins', {}).items():
        new_margin = st.number_input(
            f"Margin for weights {bracket[0]}-{bracket[1]} lbs:",
            value=default_margin,
            min_value=0,
            max_value=100,
            step=1,
            key=f'margin_{bracket[0]}_{bracket[1]}'
        )
        st.session_state.setdefault('custom_margins', {})[bracket] = new_margin  # Using setdefault() to create key if not present

    submit_custom_margins = st.form_submit_button("Submit Margins")
    if submit_custom_margins:
        st.success("Custom margins updated successfully!")

         # Add service level, date, and time of generation
        service_level = st.session_state.get('service_type', 'N/A')
        generation_date_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        st.write(f"Calculating costs for service level: {service_level}")
        st.write(f"Date & Time of Generation: {generation_date_time}")

        # Calculate Costs
        service_type = st.session_state.get('service_type')
        user_selected_zone_number = st.session_state.get('user_selected_zone_number')
        user_selected_pickup_details = firstmile_zone_details.get(user_selected_zone_number, {})
        custom_margins = st.session_state.get('custom_margins', {})  #Retrieving custom margins from session state
        if service_type and user_selected_zone_number and user_selected_pickup_details:
            target_margin_input = st.session_state.get('custom_margins', {}).get((0, 0), 0)  #Storing custom margin input in targetmargininput
            if service_type == "End to End without Pickup" or service_type == "Final Mile Only":
                    st.write("Selected Pickup Zone: NA")
            else:
                    st.write(f"Selected Pickup Zone: {st.session_state.get('pick_up_location')}")
            costs_df = calculate_costs(
                service_type,  
                user_selected_zone_number, 
                user_selected_pickup_details, 
                firstmile_variance_details,
                firstmile_zone_details, 
                final_mile_variance_rates, 
                middle_mile_pickup_details, 
                middle_mile_variance_details, 
                final_mile_costs,
                custom_margins
            )

            if costs_df is not None:
                st.session_state['costs_df'] = costs_df
            else:
                st.error("Failed to calculate costs. Please check the input details.")
        else:
            st.error("Required session state variables missing.")


# Initialize an empty dictionary to store user inputs
user_inputs_dict = {
    'Opportunity Name': [opportunity_name],
    'Quote Prepared By': [quote_prepared_by],
    'Freight Pickup Service': [freight_pickup_service],
    'Sort Initial Freight': [sort_initial_freight],
    'Pickup Location': [pick_up_location],
    'Pickup Schedule': [pick_up_schedule],
    'Pickup Timings': [pick_up_timings],
    'Average Shipment Weight': [average_shipment_weight],
    'Rate Shopping System': [use_rate_shopping_system],
    'Shipping SLA': [shipping_sla],
    'Service Type Required': [service_type_required],
    'Average Shipments Per Day': [avg_shipments_per_day],
    'Average Pieces Per Shipment': [avg_pieces_per_shipment]
}

# Convert the dictionary to a DataFrame
user_inputs_df = pd.DataFrame(user_inputs_dict)

# Display the DataFrame
#user_inputs_df

# Display calculated costs
if 'costs_df' in st.session_state and st.session_state['costs_df'] is not None and not st.session_state['costs_df'].empty:
    costs_df = st.session_state['costs_df']
    r1_df = costs_df.filter(regex='R1')
    r2_df = costs_df.filter(regex='R2')

    new_column_names = {
        'SLOK (Zone 100) - R2': 'SLOK (Zone 200) - R2',
        'SLHA/PK/BA (Zone 110) - R2': 'SLHA/PK/BA (Zone 210) - R2',
        'SLLN/WS/OR/HV (Zone 120) - R2': 'SLLN/WS/OR/HV (Zone 220) - R2',
        'SLMT (Zone 130) - R2': 'SLMT (Zone 230) - R2',
        'SLOT (Zone 135) - R2': 'SLOT (Zone 235) - R2',
        'SLQC (Zone 140) - R2': 'SLQC (Zone 240) - R2'
    }

    # Rename the columns of r2_df
    r2_df.rename(columns=new_column_names, inplace=True)

    # Display opportunity name, service level, date, and time of generation
    opportunity_name = st.session_state.get('opportunity_name', 'N/A')
    service_level = st.session_state.get('service_type', 'N/A')
    generation_datetime = datetime.datetime.now().strftime("%Y-%m-%d,%H:%M:%S")
    #st.write(f"Opportunity Name: {opportunity_name}")
    #st.write(f"Service Level: {service_level}")
    #st.write(f"Date and Time of Generation: {generation_datetime}")

    # Display R1 costs
    st.write("R1 Costs:")
    st.dataframe(r1_df.style.format("${:.2f}"))

    # Display R2 costs with renamed columns
    st.write("R2 Costs:")
    st.dataframe(r2_df.style.format("${:.2f}"))

    if st.button('Prepare R1 & R2 Ratesheet'):
        excel_data = to_excel(user_inputs_df, r1_df, r2_df)
        file_name = f"{opportunity_name}_{quote_prepared_by}_ratesheet_{generation_datetime}.xlsx"
        st.download_button(label="Download Ratesheets", data=excel_data, file_name=file_name, mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
else:
    st.write("Please calculate costs by submitting the form above.")
    st.session_state['download_ready'] = None 
