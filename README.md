# Business Intelligence Project – Data Warehouse, ETL & Dashboard
## Description

This project demonstrates the integration of data from two heterogeneous sources into a centralized data warehouse to enable business analysis and decision support.

## Project Structure
.
├── data
│   ├── scriptNorthwind.sql
│   └── .csv tables (data extracted from Access)
│
├── scripts
│   ├── etl_main.py
│   ├── Dashboard.py
│   ├── connect.py
│   ├── create_database.py
│   └── DatabaseConfig.py
│
├── reports
│   └── rapportBI.pdf
│
├── video
│   └── BIpresentation.mp4

## Prerequisites
Software

Python

SQL Server

Python Libraries

pandas

pyodbc

PyQt5

plotly

matplotlib

sqlalchemy


xlrd

matplotlib

numpy

Install the essential libraries using the following command:

pip install pandas pyodbc PyQt5 plotly

## How to Run the Project
### 1. Create the Data Warehouse

Open and run the file:

scripts/connect.py


Then run:

scripts/create_database.py

### 2. Execute the ETL Process

Run the ETL pipeline with the following command:

py scripts/etl_main.py

### 3. Launch the Dashboard

To visualize the results, start the dashboard:

py run scripts/Dashboard.py

## Author
Oulid Azouz Ahmed Chihabeddin Chafik