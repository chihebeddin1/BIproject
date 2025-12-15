# config.py
import pyodbc
import pandas as pd
from sqlalchemy import create_engine
import warnings
warnings.filterwarnings('ignore')
from connect import connect_sql_server, connect_data_werehouse

class DatabaseConfig:
    # Configuration SQL Server Northwind (source)
    SQL_SERVER = pyodbc.connect(
        'DRIVER={ODBC Driver 18 for SQL Server};'
        'SERVER=localhost;'
        'DATABASE=Northwind;'
        'Trusted_Connection=yes;'
        'Encrypt=no;'
    )
    
    # Configuration SQL Server Data Warehouse (destination)
    DW_SERVER = {
        'driver': '{ODBC Driver 18 for SQL Server}',
        'server': 'localhost;',  
        'database': 'DataWareHouse',
        'trusted_connection': 'yes',
        'Encrypt': 'no'
    }
    
    # Configuration Access (si n�cessaire)
    ACCESS_DB_PATH = r'C:\\Users\\Sos\\Desktop\\BI PROject\\Northwind 2012.accdb'  # � adapter

def create_sql_connection():
    return connect_sql_server()

def create_datawere_connection():
    return connect_data_werehouse()

def create_sqlalchemy_engine(config_dict):
    conn_str = f"mssql+pyodbc:localhost/DataWareHouse?" \
               f"driver=SQL+Server&trusted_connection=yes"
    return create_engine(conn_str) 
