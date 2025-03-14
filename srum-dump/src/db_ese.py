from pyesedb import pyesedb

def connect_to_database(file_path):
    """Connect to the ESE database and return the database object."""
    try:
        database = pyesedb.open(file_path)
        return database
    except Exception as e:
        print(f"Error connecting to database: {e}")
        return None

def query_data(database, table_name):
    """Query data from a specified table in the ESE database."""
    try:
        table = database.get_table_by_name(table_name)
        records = []
        for record in table.records:
            records.append(record)
        return records
    except Exception as e:
        print(f"Error querying data from table {table_name}: {e}")
        return []

def process_results(records):
    """Process the queried records and return formatted results."""
    processed_data = []
    for record in records:
        processed_data.append(record.get_value_data())
    return processed_data

def close_database(database):
    """Close the database connection."""
    if database:
        database.close()