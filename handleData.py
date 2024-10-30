import pandas as pd
import json
from datetime import datetime

# Cargar el archivo JSON de Instance
with open('Instance_1719906800091.json', 'r') as f:
    instance_date = json.load(f)

# Cargar el archivo JSON de Solution
with open('Solution_2024-07-02T07_54_25.241797774.json', 'r') as s:
    solution_data = json.load(s)

# Crear la tabla References desde la clave 'references'
references = instance_date['references']
df_references = pd.DataFrame(references, columns=['id', 'product'])
df_references.rename(columns={'id': 'ID_Reference', 'product': 'Product Name'}, inplace=True)

# Guardar en Excel
df_references.to_excel('References.xlsx', index=False)

# Tabla Orders

# Crear diccionario para mapear IDs a referencias desde el primer JSON
reference_mapping = {order['id']: order['reference'] for order in instance_date['orders']}

orders_data = []

# Iterar sobre los días en el archivo de solución
for day in solution_data['days']:
    # Parsear la fecha
    parsed_date = datetime.strptime(day['date'], '%Y-%m-%d')
    # Dar formato a la fecha como M/D/YYYY 8:00:00 AM
    formatted_date = f"{parsed_date.month}/{parsed_date.day}/{parsed_date.year} 8:00:00 AM"    
    # Iterar sobre las máquinas (lines) en el día actual
    for line in day['lines']:
        # Asignar un número de secuencia a cada orden procesada en esta máquina
        for creation_sequence, order in enumerate(line['orders']):
            order_id = order['id']
            reference_id = reference_mapping.get(order_id, "Referencia no encontrada")  # Obtener referencia desde el mapeo
            
            # Agregar la información a la lista
            orders_data.append({
                'ID': order_id,
                'Date': formatted_date,
                'Creation Sequence (Seconds)': creation_sequence + 1,
                'Reference': reference_id
            })

# Convertir a DataFrame y guardar en Excel
df_orders = pd.DataFrame(orders_data)
df_orders.to_excel('Orders.xlsx', index=False)

# Extraer datos para la tabla Machines desde 'lines'
machines = instance_date.get('lines', [])  # Usar el array 'lines'
if machines:
    machines_data = []
    for machine in machines:
        machine_id = machine.get('id')
        machine_type = machine.get('type', {}).get('value')  # Obtener el nombre del tipo de máquina
        
        # Concatenar la cadena con ResourceState.TotalTime(1)*3600
        absolute_time_orders = f"{machine_type}_M{machine_id}.ResourceState.TotalTime(1)*3600"
        
        machines_data.append({
            'Machine': f"{machine_type}_M{machine_id}",
            'Absolute Time Orders (seconds)': absolute_time_orders
        })
    
    df_machines = pd.DataFrame(machines_data)
    df_machines.to_excel('Machines.xlsx', index=False)
else:
    print("No se encontraron datos de Machines.")

# Extraer datos para la tabla 'Sequence Table' usando los datos del día 0
sequence_data = []
day_0 = solution_data['days'][0]  # Nos enfocamos en el primer día

for line in day_0['lines']:
    line_id = line['id']
    for order in line['orders']:
        # Crear los datos de cada columna
        sequence = f"Input@{order['type']['value']}_M{line_id}"
        order_id = order['id']
        process_time = order['finishingTime'] - order['startingTime']
        required_staff = order['requiredStaff']

        # Agregar los datos a la lista de la tabla
        sequence_data.append({
            'Sequence': sequence,
            'ID_ORDER': order_id,
            'Process Time (Seconds)': process_time,
            'Required Staff': required_staff,
            'IsSink': 0 # Indicador para mantener a las entradas de sequencia por encima de las de Sink
        })

        # Añador la fila adicional para el sink "Input@Sink1"
        sequence_data.append({
            'Sequence': "Input@Sink1",
            'ID_ORDER': order_id,
            'Process Time (Seconds)': 0,
            'Required Staff': 0,
            'IsSink': 1 # Indicador para mantener al Sink por debajo de su entrada de sequencia
        })

# Convertir a DataFrame, ordenar y guardar en Excel
df_sequence = pd.DataFrame(sequence_data)
df_sequence = df_sequence.sort_values(by=['ID_ORDER', 'IsSink']).drop(columns='IsSink')  # Sort and remove IsSink column
df_sequence.to_excel('Sequence Table.xlsx', index=False)

# Extraer datos para la tabla Setup Times desde 'lines'
setup_times_data = []
if machines:
    for machine in machines:
        for setup in machine.get('setups', []):  # 'setups' asume que existe esta lista
            setup_times_data.append({
                'Machine': f"{machine.get('type', {}).get('value')}_M{machine.get('id')}",
                'From Value': setup.get('sourceId'),  # ID de la referencia (producto) desde
                'To Value': setup.get('targetId'),  # ID de la referencia (producto) de destino
                'SetupTimes (Seconds)': setup.get('time')
            })
    
    df_setup_times = pd.DataFrame(setup_times_data)
    df_setup_times.to_excel('Setup Times.xlsx', index=False)
else:
    print("No se encontraron datos de Setup Times.")
