import pandas as pd
import numpy as np

if __name__ == '__main__':
    # Define the columns and possible values
    columns = ['id_order', 'id_leg', 'accounting_month', 'co2e_kg', 'customer', 'contractor', 'mode', 'vehicle', 'weight_kg', 'distance_km']
    id_order = np.arange(1, 21)
    id_leg = np.arange(1, 21)
    accounting_month = pd.date_range(start='1/1/2022', periods=12, freq='M')
    co2e_kg = np.random.uniform(100, 1000, 20)
    customer = ['Customer ' + str(i) for i in range(1, 6)]
    contractor = ['Contractor ' + str(i) for i in range(1, 6)]
    mode = ['Air', 'Sea', 'Rail', 'Road']
    vehicle = ['Vehicle ' + str(i) for i in range(1, 6)]
    weight_kg = np.random.uniform(1000, 20000, 20)
    distance_km = np.random.uniform(100, 1000, 20)

    # Generate random data for each column
    data = {
        'id_order': np.random.choice(id_order, 20),
        'id_leg': np.random.choice(id_leg, 20),
        'accounting_month': np.random.choice(accounting_month, 20),
        'co2e_kg': co2e_kg,
        'customer': np.random.choice(customer, 20),
        'contractor': np.random.choice(contractor, 20),
        'mode': np.random.choice(mode, 20),
        'vehicle': np.random.choice(vehicle, 20),
        'weight_kg': weight_kg,
        'distance_km': distance_km
    }

    # Create DataFrame
    df = pd.DataFrame(data, columns=columns)

    # Print DataFrame
    df.to_csv('emission_data.csv', index=False)